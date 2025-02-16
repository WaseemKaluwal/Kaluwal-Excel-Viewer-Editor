import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
import ttkbootstrap as tb
import os

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kaluwal Excel Viewer & Editor")
        self.root.geometry("1500x800")
        
        # Use ttkbootstrap for dark mode
        self.style = tb.Style("darkly")
        
        # Sidebar Frame
        self.sidebar = ttk.Frame(self.root, padding=10)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        
        fields = ["Name", "Age", "Phone", "Email", "Address", "City", "Start Date", "End Date"]
        self.entries = {}
        
        for field in fields:
            ttk.Label(self.sidebar, text=f"{field}:").pack(anchor="w")
            entry = ttk.Entry(self.sidebar)
            entry.pack(fill=tk.X, pady=5)
            self.entries[field] = entry
        
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
        
        self.edit_button = ttk.Button(self.sidebar, text="Edit Selected", command=self.edit_data)
        self.edit_button.pack(fill=tk.X, pady=5)
        
        self.load_button = ttk.Button(self.sidebar, text="Load Excel", command=self.load_excel)
        self.load_button.pack(fill=tk.X, pady=5)
        
        # Table Frame
        self.table_frame = ttk.Frame(self.root)
        self.table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        columns = fields + ["Subscription", "Employment"]
        self.tree = ttk.Treeview(self.table_frame, columns=columns, show='headings')
        
        for col in columns:
            self.tree.column(col, anchor="center", width=120, stretch=True)
            self.tree.heading(col, text=col, anchor="center")
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<ButtonRelease-1>", self.select_item)
        
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
        
        values = [self.entries[field].get() for field in self.entries]
        values.append(self.subscription_var.get())
        values.append("Employed" if self.employment_var.get() else "Unemployed")
        
        if not values[0] or not values[1].isdigit():
            messagebox.showerror("Input Error", "Please enter a valid Name and Age.")
            return
        
        try:
            wb = openpyxl.load_workbook(self.filepath)
            sheet = wb.active
            sheet.append(values)
            wb.save(self.filepath)
            wb.close()
            
            self.tree.insert("", tk.END, values=values)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to insert data: {e}")
    
    def select_item(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        
        item = self.tree.item(selected[0], "values")
        
        for i, field in enumerate(self.entries):
            self.entries[field].delete(0, tk.END)
            self.entries[field].insert(0, item[i])
        
        self.subscription_var.set(item[len(self.entries)])
        self.employment_var.set(item[len(self.entries) + 1] == "Employed")
    
    def edit_data(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an entry to edit.")
            return
        
        new_values = [self.entries[field].get() for field in self.entries]
        new_values.append(self.subscription_var.get())
        new_values.append("Employed" if self.employment_var.get() else "Unemployed")
        
        if not new_values[0] or not new_values[1].isdigit():
            messagebox.showerror("Input Error", "Please enter a valid Name and Age.")
            return
        
        try:
            wb = openpyxl.load_workbook(self.filepath)
            sheet = wb.active
            
            row_index = self.tree.index(selected[0]) + 2  # Excel rows start from 1 and we have headers
            for i, value in enumerate(new_values, start=1):
                sheet.cell(row=row_index, column=i, value=value)
            
            wb.save(self.filepath)
            wb.close()
            
            self.tree.item(selected[0], values=new_values)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update data: {e}")

if __name__ == "__main__":
    root = tb.Window(themename="darkly")
    app = ExcelApp(root)
    root.mainloop()
