import tkinter as tk
from tkinter import filedialog, messagebox
from automation import SefazAutomation
from utils import readstate_inclusions_from_excel

class AutomationUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Automation of Statement Issuance")
        
        self.label = tk.Label(self.root, text="Select the date (MM/YYYY):")
        self.label.pack(pady=10)
        
        self.entry = tk.Entry(self.root)
        self.entry.pack(pady=5)
        
        self.button = tk.Button(self.root, text="Select Excel File", command=self.select_file)
        self.button.pack(pady=10)
        
        self.run_button = tk.Button(self.root, text="Run Automation", command=self.run_automation)
        self.run_button.pack(pady=10)
        
        self.file_path = ""
    
    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            messagebox.showinfo("File Selected", f"Selected file: {self.file_path}")
    
    def run_automation(self):
        if not self.file_path:
            messagebox.showwarning("No file selected", "Please select an Excel file.")
            return
        
        period = self.entry.get()
        if not period:
            messagebox.showwarning("Date not filled", "Please enter the date (MM/YYYY).")
            return
        
        try:
            state_inclusions = readstate_inclusions_from_excel(self.file_path)
            automation = SefazAutomation()
            automation.run_automation(state_inclusions, period)
            messagebox.showinfo("Completed", "Automation completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    ui = AutomationUI()
    ui.run()
