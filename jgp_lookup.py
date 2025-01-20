import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from pathlib import Path
import os
from ttkthemes import ThemedTk

class JGPLookup:
    def __init__(self, root):
        self.root = root
        self.root.title("JGP Data Lookup")
        self.root.geometry("1200x800")
        
        # Set theme
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure styles
        self.style.configure('Header.TLabel', font=('Segoe UI', 24, 'bold'))
        self.style.configure('Subheader.TLabel', font=('Segoe UI', 12))
        self.style.configure('Search.TButton', font=('Segoe UI', 11))
        
        # Initialize data
        self.excel_data = None
        self.load_excel_data()
        
        # Create main container
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Configure grid
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        
        # Create UI elements
        self.create_header()
        self.create_search_section()
        self.create_results_section()
        
    def load_excel_data(self):
        """Load Excel data from Downloads folder"""
        downloads_path = r"C:\Users\Lenovo\OneDrive\Documents\Jiinue Growth Programme\JGP.xlsx"
        try:
            self.excel_data = pd.read_excel(downloads_path)
            # Get unique counties
            self.counties = sorted(self.excel_data['County'].unique())
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel file from Downloads folder: {e}")
            self.excel_data = None
            self.counties = []
    
    def create_header(self):
        """Create header section"""
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, pady=(0, 20), sticky="ew")
        
        title = ttk.Label(header_frame, text="JGP Data Lookup", style='Header.TLabel')
        title.grid(row=0, column=0, pady=(0, 5))
        
        subtitle = ttk.Label(header_frame, 
                           text="Easy way to find and view participant information",
                           style='Subheader.TLabel')
        subtitle.grid(row=1, column=0)
        
        header_frame.columnconfigure(0, weight=1)
    
    def create_search_section(self):
        """Create search controls"""
        search_frame = ttk.LabelFrame(self.main_frame, text="Search Criteria", padding="10")
        search_frame.grid(row=1, column=0, pady=(0, 20), sticky="ew")
        
        # County selection
        ttk.Label(search_frame, text="County:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.county_var = tk.StringVar()
        self.county_combo = ttk.Combobox(search_frame, textvariable=self.county_var)
        self.county_combo['values'] = [''] + list(self.counties)
        self.county_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # ID Number
        ttk.Label(search_frame, text="ID Number:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.id_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.id_var).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        # Phone Number
        ttk.Label(search_frame, text="Phone Number:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.phone_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.phone_var).grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        # Full Name
        ttk.Label(search_frame, text="Full Name:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.name_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.name_var).grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        
        # Search button
        search_btn = ttk.Button(search_frame, text="Search Records", 
                              command=self.search_records, style='Search.TButton')
        search_btn.grid(row=4, column=0, columnspan=2, pady=15)
        
        # Configure grid
        search_frame.columnconfigure(1, weight=1)
    
    def create_results_section(self):
        """Create results tree view"""
        # Results label
        self.results_label = ttk.Label(self.main_frame, text="Search Results")
        self.results_label.grid(row=2, column=0, pady=(0, 5), sticky="w")
        
        # Create Treeview
        columns = ('Full Name', 'ID Number', 'Phone Number', 'County', 'Gender', 
                  'Age', 'Industry', 'Business Registered')
        self.tree = ttk.Treeview(self.main_frame, columns=columns, show='headings')
        
        # Configure columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)  # Adjust width as needed
        
        # Add scrollbars
        y_scroll = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(self.main_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Grid scrollbars and tree
        self.tree.grid(row=3, column=0, sticky="nsew")
        y_scroll.grid(row=3, column=1, sticky="ns")
        x_scroll.grid(row=4, column=0, sticky="ew")
        
        # Configure grid weights
        self.main_frame.rowconfigure(3, weight=1)
    
    def search_records(self):
        """Perform search based on criteria"""
        if self.excel_data is None:
            messagebox.showerror("Error", "No data loaded")
            return
            
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get search criteria
        county = self.county_var.get()
        id_number = self.id_var.get()
        phone = self.phone_var.get()
        name = self.name_var.get()
        
        if not county:
            messagebox.showwarning("Warning", "Please select a county")
            return
            
        # Filter data
        results = self.excel_data[self.excel_data['County'].str.lower() == county.lower()]
        
        if id_number:
            results = results[results['WHAT IS YOUR NATIONAL ID?'].astype(str) == id_number]
        if phone:
            results = results[results['Phone Number'].astype(str) == phone]
        if name:
            results = results[results['Full Name'].str.lower().str.contains(name.lower())]
            
        # Update results
        for _, row in results.iterrows():
            self.tree.insert('', 'end', values=(
                row['Full Name'],
                row['WHAT IS YOUR NATIONAL ID?'],
                row['Phone Number'],
                row['County'],
                row['Gender'],
                row['Age'],
                row['WHAT IS THE MAIN INDUSTRY SECTOR IN WHICH YOU OPERATE IN?'],
                row['IS YOUR BUSINESS REGISTERED?']
            ))
            
        # Update results label
        count = len(results)
        self.results_label.configure(text=f"Search Results ({count} {'record' if count == 1 else 'records'} found)")

def main():
    root = ThemedTk(theme="clam")  # Use clam theme for a modern look
    app = JGPLookup(root)
    root.mainloop()

if __name__ == "__main__":
    main()