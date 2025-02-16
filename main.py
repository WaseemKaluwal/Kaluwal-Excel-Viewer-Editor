import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
import ttkbootstrap as tb
import os

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kaluwal Excel Viewer & Editor")
        self.root.geometry("1000x500")
        
        # Use ttkbootstrap for dark mode
        self.style = tb.Style("darkly")
        
        # Sidebar Frame
        self.sidebar = ttk.Frame(self.root, padding=10)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        
        ttk.Label(self.sidebar, text="Name:").pack(anchor="w")
        self.name_entry = ttk.Entry(self.sidebar)
        self.name_entry.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.sidebar, text="Age:").pack(anchor="w")
        self.age_entry = ttk.Entry(self.sidebar)
        self.age_entry.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.sidebar, text="Subscription:").pack(anchor="w")
        self.subscription_var = tk.StringVar(value="Subscribed")
        self.subscription_menu = ttk.Combobox(self.sidebar, textvariable=self.subscription_var,
                                              values=["Subscribed", "Unsubscribed"])
        self.subscription_menu.pack(fill=tk.X, pady=5)
        
        self.employment_var = tk.BooleanVar()
        self.employment_check = ttk.Checkbutton(self.sidebar, text="Employed", variable=self.employment_var)
        self.employment_check.pack(anchor="w", pady=5)
        
        self.insert_button = ttk.Button(self.sidebar, text="Insert", command=self.insert_data)
        self.insert_button.pack(fill=tk.X, pady=5)
        
        self.load_button = ttk.Button(self.sidebar, text="Load Excel", command=self.load_excel)
        self.load_button.pack(fill=tk.X, pady=5)
        
        # Table Frame
        self.table_frame = ttk.Frame(self.root)
        self.table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(self.table_frame, columns=("Name", "Age", "Subscription", "Employment"), show='headings')
        
        # Adjust column width and alignment
        self.tree.column("Name", anchor="center", width=150, stretch=True)
        self.tree.column("Age", anchor="center", width=80, stretch=True)
        self.tree.column("Subscription", anchor="center", width=120, stretch=True)
        self.tree.column("Employment", anchor="center", width=120, stretch=True)
        
        self.tree.heading("Name", text="Name", anchor="center")
        self.tree.heading("Age", text="Age", anchor="center")
        self.tree.heading("Subscription", text="Subscription", anchor="center")
        self.tree.heading("Employment", text="Employment", anchor="center")
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        self.filepath = None

    def load_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not filepath:
            return
        
        self.filepath = filepath
        self.tree.delete(*self.tree.get_children())
        
        try:
            wb = openpyxl.load_workbook(filepath)
            sheet = wb.active
            
            for row in sheet.iter_rows(min_row=2, values_only=True):
                self.tree.insert("", tk.END, values=row)
            
            wb.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
    
    def insert_data(self):
        if not self.filepath:
            messagebox.showwarning("Warning", "Please load an Excel file first.")
            return
        
        name = self.name_entry.get()
        age = self.age_entry.get()
        subscription = self.subscription_var.get()
        employment = "Employed" if self.employment_var.get() else "Unemployed"
        
        if not name or not age.isdigit():
            messagebox.showerror("Input Error", "Please enter a valid Name and Age.")
            return
        
        try:
            wb = openpyxl.load_workbook(self.filepath)
            sheet = wb.active
            sheet.append([name, int(age), subscription, employment])
            wb.save(self.filepath)
            wb.close()
            
            self.tree.insert("", tk.END, values=(name, age, subscription, employment))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to insert data: {e}")

if __name__ == "__main__":
    root = tb.Window(themename="darkly")
    app = ExcelApp(root)
    root.mainloop()
