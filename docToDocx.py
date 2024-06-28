import os
import time
from tkinter import Tk, Button, Text, Scrollbar, filedialog, ttk
import win32com.client as win32
from threading import Thread

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DOC to DOCX Converter")
        self.root.geometry("500x400")  # Fixed size: width=500, height=400

        self.textbox = Text(self.root, wrap='word')
        self.textbox.pack(expand=True, fill='both')

        self.scrollbar = Scrollbar(self.textbox)
        self.scrollbar.config(command=self.textbox.yview)
        self.textbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side='right', fill='y')

        self.progress = ttk.Progressbar(self.root, orient='horizontal', mode='determinate')

        self.estimated_time_label = Text(self.root, height=1, wrap='word')
        self.estimated_time_label.insert('1.0', "Estimated time remaining: Calculating...")

        self.processed_files_label = Text(self.root, height=1, wrap='word')
        self.processed_files_label.insert('1.0', "N° of processed files :")

        self.convert_button = Button(self.root, text="Select Folder and Convert", command=self.convert_folder)
        self.convert_button.pack(pady=10)

        self.stop_button = Button(self.root, text="Stop Conversion", command=self.stop_conversion, state='disabled')

        self.close_button = Button(self.root, text="Close", command=self.close_app)
        self.close_button.pack(pady=10)

        self.conversion_in_progress = False
        self.stop_conversion_flag = False

    def convert_doc_to_docx(self, word, doc_files):
        try:
            for doc_file in doc_files:
                if self.stop_conversion_flag:
                    self.update_textbox("Conversion stopped by user.")
                    return False

                docx_file = os.path.join(os.path.dirname(doc_file), f"{os.path.splitext(os.path.basename(doc_file))[0]}.docx")
                if not os.path.exists(docx_file):
                    doc = word.Documents.Open(doc_file)
                    doc.SaveAs(docx_file, FileFormat=16)  # 16 represents the docx format
                    doc.Close()
                    folder_path = os.path.dirname(docx_file)
                    self.update_textbox(f"Converted: {folder_path}\{os.path.basename(doc_file)}")

                if self.stop_conversion_flag:
                    self.update_textbox("Conversion stopped by user.")
                    return False

            return True
        except Exception as e:
            print(f"Error converting files: {str(e)}")
            return False

    def convert_folder_recursively(self, folder_path):
        self.convert_button.pack_forget()
        self.close_button.pack_forget()

        doc_files = []
        for dirpath, _, filenames in os.walk(folder_path):
            for filename in filenames:
                if filename.endswith('.doc'):
                    doc_files.append(os.path.join(dirpath, filename))

        total_files = len(doc_files)
        chunk_size = 10
        processed_files = 0

        start_time = time.time()

        self.conversion_in_progress = True
        self.stop_button.pack(pady=10)  # Show stop button


        # Batch process DOC files
        word = win32.Dispatch('Word.Application')
        while processed_files < total_files:
            if self.stop_conversion_flag:
                break

            chunk_files = doc_files[processed_files:processed_files + chunk_size]
            if self.convert_doc_to_docx(word, chunk_files):
                processed_files += len(chunk_files)
                self.update_progress(processed_files, total_files)

                elapsed_time = time.time() - start_time
                estimated_total_time = (elapsed_time / processed_files) * total_files
                estimated_remaining_time = estimated_total_time - elapsed_time
                self.update_estimated_time(estimated_total_time)
                self.processed_files_label.delete('1.0', 'end')
                self.processed_files_label.insert('1.0', f'N° of processed files :{processed_files} of {total_files}')

        word.Quit()
        self.conversion_in_progress = False
        self.stop_button.pack_forget()  # Hide stop button after conversion
        self.progress.pack_forget()  # Hide progress bar after conversion
        self.estimated_time_label.pack_forget()  # Hide estimated time label after conversion
        self.processed_files_label.pack_forget()
        self.convert_button.pack(pady=10)
        self.close_button.pack(pady=10)

    def stop_conversion(self):
        if self.conversion_in_progress:
            self.stop_conversion_flag = True
            self.update_textbox("Stopping conversion...")

    def update_textbox(self, message):
        self.textbox.insert('end', message + '\n')
        self.textbox.see('end')
        self.root.update_idletasks()

    def update_progress(self, processed, total):
        self.progress['value'] = (processed / total) * 100
        self.root.update_idletasks()

    def update_estimated_time(self, remaining_time):
        minutes, seconds = divmod(int(remaining_time), 60)
        time_str = f"Estimated time remaining: {minutes} min {seconds} sec"
        self.estimated_time_label.delete('1.0', 'end')
        self.estimated_time_label.insert('1.0', time_str)
        self.root.update_idletasks()

    def get_folder_path(self):
        try:
            folder_path = filedialog.askdirectory(title="Select Folder")
            if folder_path:
                folder_path = os.path.abspath(folder_path)
                return folder_path
            else:
                return None
        except Exception as e:
            print(f"Error retrieving folder path: {e}")
            return None

    def convert_folder(self):
        if not self.conversion_in_progress:
            folder_path = self.get_folder_path()
            if folder_path:
                self.progress.pack(fill='x', pady=10)  # Show progress bar
                self.processed_files_label.pack(fill='x', pady=10)
                self.estimated_time_label.pack(fill='x', pady=10)  # Show estimated time label
                self.update_textbox(f"Converting files in folder: {folder_path}")
                self.stop_button.config(state='normal')  # Enable stop button
                self.stop_conversion_flag = False
                Thread(target=self.convert_folder_recursively, args=(folder_path,)).start()

    def close_app(self):
        self.root.destroy()

if __name__ == "__main__":
    root = Tk()
    app = FileConverterApp(root)
    root.mainloop()
