import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import boto3
import json
import os
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class TranslatorApp:
    def __init__(self, master):
        self.master = master
        self.df = None  # 初始化 self.df 为 None
        master.title("Excel/CSV Translator")
        master.geometry("600x500")

        logger.info("Initializing TranslatorApp")

        # 创建主框架
        main_frame = ttk.Frame(master, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        # 源语言和目标语言选择
        self.source_lang = tk.StringVar(value="English")
        self.target_lang = tk.StringVar(value="Chinese")

        ttk.Label(main_frame, text="Source Language:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Combobox(main_frame, textvariable=self.source_lang, values=["English", "Chinese"], state="readonly").grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(main_frame, text="Target Language:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Combobox(main_frame, textvariable=self.target_lang, values=["English", "Chinese"], state="readonly").grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)

        # 文件选择
        self.file_path = tk.StringVar()
        ttk.Button(main_frame, text="Select File", command=self.select_file).grid(row=2, column=0, columnspan=2, pady=10)
        self.file_label = ttk.Label(main_frame, text="No file selected")
        self.file_label.grid(row=3, column=0, columnspan=2, pady=5)

        # Sheet选择 (for xlsx)
        self.sheet_frame = ttk.Frame(main_frame)
        self.sheet_frame.grid(row=4, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        ttk.Label(self.sheet_frame, text="Select Sheet:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(self.sheet_frame, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # 列选择
        ttk.Label(main_frame, text="Select Columns to Translate:").grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)
        self.columns_listbox = tk.Listbox(main_frame, selectmode=tk.MULTIPLE, exportselection=0)
        self.columns_listbox.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(6, weight=1)
        
        self.use_new_columns_var = tk.BooleanVar()
        self.use_new_columns_checkbox = ttk.Checkbutton(
            main_frame,  # 替换为您实际的父容器
            text="使用新列存放翻译内容",
            variable=self.use_new_columns_var
        )
        self.use_new_columns_checkbox.grid(row=5, column=0, sticky="w")  


        # 翻译按钮
        ttk.Button(main_frame, text="Translate", command=self.translate).grid(row=7, column=0, columnspan=2, pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if file_path:
            logger.info(f"Selected file: {file_path}")
            self.file_path.set(file_path)
            self.file_label.config(text=os.path.basename(file_path))
            self.load_file(file_path)

    def load_file(self, file_path):
        logger.info(f"Loading file: {file_path}")
        if file_path.endswith('.xlsx'):
            self.df = pd.read_excel(file_path, sheet_name=None)
            self.sheet_frame.grid()
            self.sheet_combo['values'] = list(self.df.keys())
            self.sheet_combo.set(list(self.df.keys())[0])
            self.sheet_combo.bind('<<ComboboxSelected>>', self.update_columns)
            self.update_columns()
            logger.info(f"Loaded Excel file with sheets: {list(self.df.keys())}")
        elif file_path.endswith('.csv'):
            self.df = pd.read_csv(file_path)
            self.sheet_frame.grid_remove()
            self.update_columns_csv()
            logger.info("Loaded CSV file")

    def update_columns(self, event=None):
        selected_sheet = self.sheet_combo.get()
        columns = list(self.df[selected_sheet].columns)
        self.columns_listbox.delete(0, tk.END)
        for col in columns:
            self.columns_listbox.insert(tk.END, col)
        logger.info(f"Updated columns for sheet: {selected_sheet}")

    def update_columns_csv(self):
        columns = list(self.df.columns)
        self.columns_listbox.delete(0, tk.END)
        for col in columns:
            self.columns_listbox.insert(tk.END, col)
        logger.info("Updated columns for CSV file")

    def translate(self):
        if not hasattr(self, 'df') or self.df is None:
            logger.error("No file selected")
            messagebox.showerror("Error", "Please select a file first.")
            return

        selected_columns = [self.columns_listbox.get(i) for i in self.columns_listbox.curselection()]
        if not selected_columns:
            logger.error("No columns selected for translation")
            messagebox.showerror("Error", "Please select at least one column to translate.")
            return

        source_lang = self.source_lang.get()
        target_lang = self.target_lang.get()
        logger.info(f"Translating from {source_lang} to {target_lang}")

        use_new_columns = self.use_new_columns_var.get()

        if self.file_path.get().endswith('.xlsx'):
            selected_sheet = self.sheet_combo.get()
            df_to_translate = self.df[selected_sheet]
            logger.info(f"Translating sheet: {selected_sheet}")
        else:
            df_to_translate = self.df
            logger.info("Translating CSV file")

        translated_df = self.translate_dataframe(df_to_translate, selected_columns, source_lang, target_lang, use_new_columns)

        # Automatically save the translated file
        input_file_path = self.file_path.get()
        input_file_name = os.path.basename(input_file_path)
        input_file_dir = os.path.dirname(input_file_path)
        file_name, file_extension = os.path.splitext(input_file_name)
        
        output_file_name = f"{file_name}_translated{file_extension}"
        output_path = os.path.join(input_file_dir, output_file_name)

        # Ensure we don't overwrite an existing file
        counter = 1
        while os.path.exists(output_path):
            output_file_name = f"{file_name}_translated_{counter}{file_extension}"
            output_path = os.path.join(input_file_dir, output_file_name)
            counter += 1

        if file_extension.lower() == '.xlsx':
            translated_df.to_excel(output_path, index=False,encoding='utf-8-sig')
        else:  # Assume CSV
            translated_df.to_csv(output_path, index=False,encoding='utf-8-sig')

        logger.info(f"Translated file saved as {output_path}")
        messagebox.showinfo("Success", f"Translated file saved as {output_path}")
        
    def translate_dataframe(self, df, columns, source_lang, target_lang, use_new_columns):
        bedrock = boto3.client('bedrock-runtime', region_name='us-west-2')  # 替换为你的区域
        logger.info("Initializing Bedrock client")

        model_id = "anthropic.claude-3-haiku-20240307-v1:0"

        for column in columns:
            logger.info(f"Translating column: {column}")
            translated_column = []
            for index, text in enumerate(df[column]):
                if pd.notna(text):
                    prompt = f"Translate the following text from {source_lang} to {target_lang}. Only provide the translated text without any additional explanations or the original text such as 'here is the ...' : {text}"
                    
                    native_request = {
                        "anthropic_version": "bedrock-2023-05-31",
                        "max_tokens": 1000,
                        "temperature": 0.1,
                        "messages": [
                            {
                                "role": "user",
                                "content": [{"type": "text", "text": prompt}],
                            }
                        ],
                    }

                    request = json.dumps(native_request)

                    try:
                        response = bedrock.invoke_model(modelId=model_id, body=request)
                        model_response = json.loads(response["body"].read())
                        
                        logger.info(f"API response: {model_response}")  # 记录完整的 API 响应
                        
                        if 'content' in model_response and model_response['content']:
                            translated_text = model_response['content'][0]['text'].strip()
                            logger.info(f"Translated text {index + 1}/{len(df[column])} in column {column}")
                        else:
                            logger.error(f"Unexpected API response for text {index + 1} in column {column}: {model_response}")
                            translated_text = "API response error"
                    except Exception as e:
                        logger.error(f"Translation error for text {index + 1} in column {column}: {str(e)}")
                        logger.error(f"API response: {model_response if 'model_response' in locals() else 'No response'}")
                        translated_text = "Translation error"
                else:
                    translated_text = ''
                
                translated_column.append(translated_text)
            
            logger.info(f"Length of translated column: {len(translated_column)}")
            logger.info(f"Length of original DataFrame: {len(df)}")
            
            if len(translated_column) != len(df):
                logger.error(f"Mismatch in lengths for column {column}. Filling missing values with 'Translation error'")
                translated_column.extend(['Translation error'] * (len(df) - len(translated_column)))
            
            if use_new_columns:
                # 使用新列存放翻译内容
                new_column_name = f"{column}_translated"
                df[new_column_name] = translated_column
                logger.info(f"Created new column: {new_column_name}")
            else:
                # 在原列上覆盖翻译内容
                df[column] = translated_column
                logger.info(f"Updated original column: {column}")

        return df


if __name__ == "__main__":
    logger.info("Starting TranslatorApp")
    root = tk.Tk()
    app = TranslatorApp(root)
    root.mainloop()
    logger.info("TranslatorApp closed")
