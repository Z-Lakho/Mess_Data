# import tkinter as tk
# from tkinter import ttk, filedialog, messagebox
# import pandas as pd
# from datetime import datetime
# import os
# from utils import get_mess_charges
# import sys

# class MessManagementApp(tk.Tk):
#     def __init__(self):
#         super().__init__()
#         self.title("Mess Management System")
#         self.geometry("1200x600")
#         self.configure(bg="#f0f4f8")

#         # Handle paths for frozen vs. non-frozen environments
#         if getattr(sys, 'frozen', False):
#             # Running as a PyInstaller executable
#             base_path = sys._MEIPASS
#         else:
#             # Running as a regular Python script
#             base_path = os.path.dirname(__file__)
        
#         self.data_dir = os.path.join(base_path, "data")
#         if not os.path.exists(self.data_dir):
#             os.makedirs(self.data_dir)
#         self.mess_file = os.path.join(self.data_dir, "mess_data.xlsx")
#         self.df_mess = None

#         if os.path.exists(self.mess_file):
#             self.df_mess = pd.read_excel(self.mess_file)
#             self.df_mess['S#'] = pd.to_numeric(self.df_mess['S#'], errors='coerce').fillna(0).astype(int)
#             self.df_mess['Mess Charges'] = pd.to_numeric(self.df_mess['Mess Charges'], errors='coerce').fillna(0).astype(int)
#             self.df_mess['SalaryMonth'] = self.df_mess['SalaryMonth'].astype(str)
#         else:
#             self.upload_initial_mess()

#         self.style = ttk.Style()
#         self.style.configure("TButton", font=("Helvetica", 10), padding=10, foreground="black")
#         self.style.configure("TLabel", font=("Helvetica", 10), background="#f0f4f8", foreground="black")
#         self.style.configure("Treeview", font=("Helvetica", 9), rowheight=25, foreground="black")
#         self.style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"), foreground="black")
#         self.style.map("TButton", background=[("active", "#0052cc"), ("!active", "#007bff")],
#                        foreground=[("active", "black"), ("!active", "black")])
#         self.style.configure("TCombobox", font=("Helvetica", 10), foreground="black")

#         self.create_widgets()
#         self.display_data()  # Load data on startup

#     def upload_initial_mess(self):
#         file_path = filedialog.askopenfilename(title="Upload Initial Mess Data", filetypes=[("Excel files", "*.xlsx *.xls")])
#         if file_path:
#             self.df_mess = pd.read_excel(file_path)
#             self.df_mess['S#'] = pd.to_numeric(self.df_mess['S#'], errors='coerce').fillna(0).astype(int)
#             self.df_mess['Mess Charges'] = pd.to_numeric(self.df_mess['Mess Charges'], errors='coerce').fillna(0).astype(int)
#             self.df_mess['SalaryMonth'] = self.df_mess['SalaryMonth'].astype(str)
#             self.df_mess.to_excel(self.mess_file, index=False)
#             messagebox.showinfo("Success", "Initial mess data uploaded and copied.")
#             self.create_widgets()
#             self.display_data()
#         else:
#             messagebox.showerror("Error", "No file selected. Exiting.")
#             self.quit()

#     def create_widgets(self):
#         main_frame = tk.Frame(self, bg="#f0f4f8")
#         main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

#         button_frame = tk.Frame(main_frame, bg="#f0f4f8")
#         button_frame.pack(fill=tk.X, pady=10)

#         upload_master_btn = ttk.Button(button_frame, text="Upload Monthly Master Data", command=self.upload_master)
#         upload_master_btn.pack(side=tk.LEFT, padx=5)

#         # fetch_btn = ttk.Button(button_frame, text="Fetch Mess Data", command=self.display_data)
#         # fetch_btn.pack(side=tk.LEFT, padx=5)

#         search_frame = tk.Frame(main_frame, bg="#f0f4f8")
#         search_frame.pack(fill=tk.X, pady=10)

#         ttk.Label(search_frame, text="Filter by Month:").pack(side=tk.LEFT, padx=5)
#         months = sorted(self.df_mess['SalaryMonth'].unique()) if self.df_mess is not None else []
#         self.month_combo = ttk.Combobox(search_frame, values=months, width=20)
#         self.month_combo.pack(side=tk.LEFT, padx=5)
#         self.month_combo.bind("<<ComboboxSelected>>", self.filter_by_month)

#         ttk.Label(search_frame, text="Search by ID or Name:").pack(side=tk.LEFT, padx=5)
#         self.search_entry = ttk.Entry(search_frame, width=30, font=("Helvetica", 10))
#         self.search_entry.pack(side=tk.LEFT, padx=5)

#         search_btn = ttk.Button(search_frame, text="Search", command=self.search_data)
#         search_btn.pack(side=tk.LEFT, padx=5)

#         # show_all_btn = ttk.Button(search_frame, text="Show All", command=self.display_data)
#         # show_all_btn.pack(side=tk.LEFT, padx=5)

#         tree_frame = tk.Frame(main_frame, bg="#f0f4f8")
#         tree_frame.pack(fill=tk.BOTH, expand=True, pady=10)

#         self.tree = ttk.Treeview(tree_frame, columns=list(self.df_mess.columns), show="headings", style="Treeview")
#         for col in self.df_mess.columns:
#             self.tree.heading(col, text=col)
#             self.tree.column(col, width=120, anchor="center")
#         self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

#         scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
#         scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
#         self.tree.configure(yscrollcommand=scrollbar.set)

#         self.tree.tag_configure("oddrow", background="#e6f3ff", foreground="black")
#         self.tree.tag_configure("evenrow", background="#ffffff", foreground="black")

#         self.total_label = ttk.Label(main_frame, text="Total Mess Amount: 0", font=("Helvetica", 12, "bold"))
#         self.total_label.pack(pady=10)

#         self.tree.bind("<Double-1>", self.on_double_click)

#     def filter_by_month(self, event):
#         month = self.month_combo.get()
#         if month:
#             filtered = self.df_mess[self.df_mess['SalaryMonth'] == month]
#             self.display_data(filtered)
#         else:
#             self.display_data()

#     def on_double_click(self, event):
#         self.edit_selected()

#     def upload_master(self):
#         file_path = filedialog.askopenfilename(title="Upload Master Data", filetypes=[("Excel files", "*.xlsx *.xls")])
#         if file_path:
#             df_master = pd.read_excel(file_path)
#             if 'EmployeeCode' in df_master.columns:
#                 df_master = df_master.rename(columns={'EmployeeCode': 'Employee Code'})
#             current_month = datetime.now().strftime("%B %Y")
#             new_rows = []
#             max_s = self.df_mess['S#'].max() if not self.df_mess.empty else 0
#             for idx, row in df_master.iterrows():
#                 code = row['Employee Code']
#                 if not ((self.df_mess['Employee Code'] == code) & (self.df_mess['SalaryMonth'] == current_month)).any():
#                     max_s += 1
#                     name = f"{row.get('FirstName', '')} {row.get('LastName', '')}".strip()
#                     if not name:
#                         name = row.get('Employee Name', 'Unknown')
#                     status = row.get('EmployeeStatus', 'Active')
#                     station = row.get('Station', 'PMTF')
#                     department = row.get('Department', '')
#                     group = row.get('EmployeeGroup', 'Officer')
#                     designation = row.get('Designation', '')
#                     charges = get_mess_charges(group, designation)
#                     new_row = {
#                         'S#': max_s,
#                         'SalaryMonth': current_month,
#                         'Employee Code': code,
#                         'Employee Name': name,
#                         'Status': status,
#                         'StationName': station,
#                         'DepartmentName': department,
#                         'EmployeeGroup': group,
#                         'Designation': designation,
#                         'Mess Charges': charges
#                     }
#                     new_rows.append(new_row)
#             if new_rows:
#                 new_df = pd.DataFrame(new_rows)
#                 self.df_mess = pd.concat([self.df_mess, new_df], ignore_index=True)
#                 self.save_mess_data()
#                 self.month_combo['values'] = sorted(self.df_mess['SalaryMonth'].unique())
#                 messagebox.showinfo("Success", f"Added {len(new_rows)} rows for {current_month}.")
#             else:
#                 messagebox.showinfo("Info", "No new entries for this month.")
#             self.display_data()

#     def display_data(self, df=None):
#         if df is None:
#             df = self.df_mess
#         self.tree.delete(*self.tree.get_children())
#         for idx, row in df.iterrows():
#             tag = "evenrow" if idx % 2 == 0 else "oddrow"
#             self.tree.insert("", "end", values=tuple(row), tags=(tag,))
#         total = df['Mess Charges'].sum() if not df.empty else 0
#         self.total_label.config(text=f"Total Mess Amount: {total}")

#     def search_data(self):
#         query = self.search_entry.get().strip().lower()
#         if not query:
#             self.display_data()
#             return
#         filtered = self.df_mess[
#             self.df_mess['Employee Code'].astype(str).str.lower().str.contains(query) |
#             self.df_mess['Employee Name'].str.lower().str.contains(query)
#         ]
#         month = self.month_combo.get()
#         if month:
#             filtered = filtered[filtered['SalaryMonth'] == month]
#         self.display_data(filtered)

#     def delete_selected(self):
#         selected = self.tree.selection()
#         if not selected:
#             messagebox.showwarning("Warning", "Select a row to delete.")
#             return
#         confirm = messagebox.askyesno("Confirm", "Delete selected rows?")
#         if confirm:
#             for item in selected:
#                 values = self.tree.item(item)['values']
#                 code = values[2]  # Employee Code
#                 month = values[1]  # SalaryMonth
#                 self.df_mess = self.df_mess[~((self.df_mess['Employee Code'] == code) & (self.df_mess['SalaryMonth'] == month))]
#             self.df_mess['S#'] = range(1, len(self.df_mess) + 1)
#             self.save_mess_data()
#             self.month_combo['values'] = sorted(self.df_mess['SalaryMonth'].unique())
#             self.display_data()

#     def edit_selected(self):
#         selected = self.tree.selection()
#         if not selected or len(selected) != 1:
#             messagebox.showwarning("Warning", "Select one row to edit.")
#             return
#         item = selected[0]
#         values = self.tree.item(item)['values']
#         edit_win = tk.Toplevel(self)
#         edit_win.title("Edit Employee")
#         edit_win.configure(bg="#f0f4f8")

#         entries = {}
#         for i, col in enumerate(self.df_mess.columns):
#             ttk.Label(edit_win, text=col).grid(row=i, column=0, sticky="w", padx=10, pady=5)
#             entry = ttk.Entry(edit_win, width=50, font=("Helvetica", 10))
#             entry.insert(0, str(values[i]) if values[i] is not None else "")
#             entry.grid(row=i, column=1, sticky="w", padx=10, pady=5)
#             entries[col] = entry

#         def save_edit():
#             new_values = {col: entries[col].get().strip() for col in self.df_mess.columns}
#             original_code = values[2]  # Original Employee Code
#             original_month = values[1]  # Original SalaryMonth
            
#             mask = (self.df_mess['Employee Code'] == original_code) & (self.df_mess['SalaryMonth'] == original_month)
#             if not mask.any():
#                 messagebox.showerror("Error", f"Entry for {original_code} in {original_month} not found.")
#                 edit_win.destroy()
#                 return
#             idx = self.df_mess[mask].index[0]
            
#             try:
#                 for col, val in new_values.items():
#                     if not val and col not in ['Employee Name', 'DepartmentName', 'Designation']:
#                         messagebox.showerror("Error", f"{col} cannot be empty.")
#                         return
#                     if col in ['S#', 'Mess Charges']:
#                         self.df_mess.at[idx, col] = int(float(val))
#                     else:
#                         self.df_mess.at[idx, col] = val
#             except ValueError as e:
#                 messagebox.showerror("Error", f"Invalid value: {str(e)}")
#                 return
                
#             self.save_mess_data()
#             self.display_data()
#             self.month_combo['values'] = sorted(self.df_mess['SalaryMonth'].unique())
#             messagebox.showinfo("Success", "Employee data updated.")
#             edit_win.destroy()

#         ttk.Button(edit_win, text="Save", command=save_edit).grid(row=len(self.df_mess.columns), column=1, pady=20)
#         ttk.Button(edit_win, text="Delete", command=lambda: [self.delete_selected(), edit_win.destroy()]).grid(row=len(self.df_mess.columns), column=0, pady=20)

#     def save_mess_data(self):
#         self.df_mess.to_excel(self.mess_file, index=False)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
import os
import sys
import win32api
import win32print

class MessManagementApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mess Management System")
        self.geometry("1200x600")
        self.configure(bg="#f0f4f8")

        # Handle paths for frozen vs. non-frozen environments
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
            self.data_dir = os.path.expanduser("~/MessManagementData")
            try:
                if not os.path.exists(self.data_dir):
                    os.makedirs(self.data_dir)
            except Exception as e:
                print(f"Warning: Cannot create {self.data_dir}: {str(e)}. Falling back to executable directory.")
                self.data_dir = os.path.dirname(sys.executable)
                if not os.path.exists(self.data_dir):
                    os.makedirs(self.data_dir)
        else:
            base_path = os.path.dirname(__file__)
            self.data_dir = os.path.join(base_path, "data")
        
        try:
            if not os.path.exists(self.data_dir):
                os.makedirs(self.data_dir)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create data directory {self.data_dir}: {str(e)}")
            self.quit()
        
        self.mess_file = os.path.join(self.data_dir, "mess_data.xlsx")
        if not os.path.exists(self.mess_file) and getattr(sys, 'frozen', False):
            bundled_mess_file = os.path.join(base_path, "data", "mess_data.xlsx")
            try:
                if os.path.exists(bundled_mess_file):
                    import shutil
                    shutil.copy(bundled_mess_file, self.mess_file)
                else:
                    messagebox.showerror("Error", f"Bundled mess_data.xlsx not found at {bundled_mess_file}")
                    self.quit()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy mess_data.xlsx to {self.mess_file}: {str(e)}")
                self.quit()
        
        print(f"Using mess_data.xlsx at: {self.mess_file}")
        
        try:
            if os.path.exists(self.mess_file):
                self.df_mess = pd.read_excel(self.mess_file)
                self.df_mess['S#'] = pd.to_numeric(self.df_mess['S#'], errors='coerce').fillna(0).astype(int)
                self.df_mess['Mess Charges'] = pd.to_numeric(self.df_mess['Mess Charges'], errors='coerce').fillna(0).astype(int)
                self.df_mess['SalaryMonth'] = self.df_mess['SalaryMonth'].astype(str)
                # Remove duplicates based on Employee Code and SalaryMonth
                duplicates = self.df_mess[self.df_mess.duplicated(subset=['Employee Code', 'SalaryMonth'], keep=False)]
                if not duplicates.empty:
                    print(f"Found {len(duplicates)} duplicate entries: {duplicates[['Employee Code', 'SalaryMonth', 'Mess Charges']].to_dict('records')}")
                    messagebox.showwarning("Warning", f"Found {len(duplicates)} duplicate entries. Keeping the last entry for each Employee Code and SalaryMonth.")
                self.df_mess = self.df_mess.drop_duplicates(subset=['Employee Code', 'SalaryMonth'], keep='last')
                self.save_mess_data()
            else:
                self.upload_initial_mess()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load mess_data.xlsx: {str(e)}")
            self.quit()

        self.style = ttk.Style()
        self.style.configure("TButton", font=("Helvetica", 10), padding=10, foreground="black")
        self.style.configure("TLabel", font=("Helvetica", 10), background="#f0f4f8", foreground="black")
        self.style.configure("Treeview", font=("Helvetica", 9), rowheight=25, foreground="black")
        self.style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"), foreground="black")
        self.style.map("TButton", background=[("active", "#0052cc"), ("!active", "#007bff")],
                       foreground=[("active", "black"), ("!active", "black")])
        self.style.configure("TCombobox", font=("Helvetica", 10), foreground="black")

        self.current_df = None
        self.create_widgets()
        self.display_data()

    def upload_initial_mess(self):
        file_path = filedialog.askopenfilename(title="Upload Initial Mess Data", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.df_mess = pd.read_excel(file_path)
                self.df_mess['S#'] = pd.to_numeric(self.df_mess['S#'], errors='coerce').fillna(0).astype(int)
                self.df_mess['Mess Charges'] = pd.to_numeric(self.df_mess['Mess Charges'], errors='coerce').fillna(0).astype(int)
                self.df_mess['SalaryMonth'] = self.df_mess['SalaryMonth'].astype(str)
                self.df_mess = self.df_mess.drop_duplicates(subset=['Employee Code', 'SalaryMonth'], keep='last')
                self.save_mess_data()
                messagebox.showinfo("Success", "Initial mess data uploaded and copied.")
                self.create_widgets()
                self.display_data()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load initial mess data: {str(e)}")
                self.quit()
        else:
            messagebox.showerror("Error", "No file selected. Exiting.")
            self.quit()

    def create_widgets(self):
        main_frame = tk.Frame(self, bg="#f0f4f8")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        button_frame = tk.Frame(main_frame, bg="#f0f4f8")
        button_frame.pack(fill=tk.X, pady=10)

        upload_master_btn = ttk.Button(button_frame, text="Upload Monthly Master Data", command=self.upload_master)
        upload_master_btn.pack(side=tk.LEFT, padx=5)

        print_btn = ttk.Button(button_frame, text="Print Data", command=self.print_data)
        print_btn.pack(side=tk.LEFT, padx=5)

        search_frame = tk.Frame(main_frame, bg="#f0f4f8")
        search_frame.pack(fill=tk.X, pady=10)

        ttk.Label(search_frame, text="Filter by Month:").pack(side=tk.LEFT, padx=5)
        months = sorted(self.df_mess['SalaryMonth'].unique()) if self.df_mess is not None else []
        self.month_combo = ttk.Combobox(search_frame, values=months, width=20)
        self.month_combo.pack(side=tk.LEFT, padx=5)
        self.month_combo.bind("<<ComboboxSelected>>", self.filter_by_month)

        ttk.Label(search_frame, text="Search by ID or Name:").pack(side=tk.LEFT, padx=5)
        self.search_entry = ttk.Entry(search_frame, width=30, font=("Helvetica", 10))
        self.search_entry.pack(side=tk.LEFT, padx=5)

        search_btn = ttk.Button(search_frame, text="Search", command=self.search_data)
        search_btn.pack(side=tk.LEFT, padx=5)

        tree_frame = tk.Frame(main_frame, bg="#f0f4f8")
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        xscrollbar = ttk.Scrollbar(tree_frame, orient="horizontal")
        xscrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        yscrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(tree_frame, columns=list(self.df_mess.columns), show="headings", style="Treeview", xscrollcommand=xscrollbar.set, yscrollcommand=yscrollbar.set)
        xscrollbar.configure(command=self.tree.xview)
        yscrollbar.configure(command=self.tree.yview)

        for col in self.df_mess.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="w", stretch=tk.YES)  # Increased width, anchor left, stretch yes for auto-adjust

        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.tag_configure("oddrow", background="#e6f3ff", foreground="black")
        self.tree.tag_configure("evenrow", background="#ffffff", foreground="black")

        self.total_label = ttk.Label(main_frame, text="Total Mess Amount: 0", font=("Helvetica", 12, "bold"))
        self.total_label.pack(pady=10)

        self.tree.bind("<Double-1>", self.on_double_click)

    def filter_by_month(self, event):
        month = self.month_combo.get()
        if month:
            filtered = self.df_mess[self.df_mess['SalaryMonth'] == month]
            print(f"Filtering for month: {month}, {len(filtered)} rows found")
            if filtered.empty:
                messagebox.showinfo("Info", f"No data found for {month}")
            self.display_data(filtered)
        else:
            print("No month filter applied, showing all data")
            self.display_data()

    def on_double_click(self, event):
        self.edit_selected()

    def upload_master(self):
        file_path = filedialog.askopenfilename(title="Upload Master Data", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                print(f"Uploading file: {file_path}")
                df_master = pd.read_excel(file_path)
                print(f"Columns in uploaded file: {df_master.columns.tolist()}")

                normalized_columns = {col.strip().lower(): col for col in df_master.columns}
                possible_code_columns = ['employee code', 'employeecode', 'empcode', 'employee id', 'empid', 'id', 'code']
                code_column = None
                for col in normalized_columns:
                    if col in possible_code_columns:
                        code_column = normalized_columns[col]
                        break
                
                if code_column is None:
                    messagebox.showerror("Error", "No valid employee code column found. Expected one of: " + ", ".join(possible_code_columns))
                    return

                column_mapping = {
                    code_column: 'Employee Code',
                    normalized_columns.get('employee name', 'Employee Name'): 'Employee Name',
                    normalized_columns.get('employeegroup', 'EmployeeGroup'): 'EmployeeGroup',
                    normalized_columns.get('departmentname', 'DepartmentName'): 'DepartmentName',
                    normalized_columns.get('designationname', 'DesignationName'): 'Designation',
                    normalized_columns.get('messcharges', 'MessCharges'): 'Mess Charges',
                    normalized_columns.get('payrollstartdate', 'PayrollStartDate'): 'PayrollStartDate'
                }
                df_master = df_master.rename(columns=column_mapping)

                required_columns = ['Employee Code', 'Employee Name', 'EmployeeGroup', 'DepartmentName', 'Designation', 'Mess Charges', 'PayrollStartDate']
                missing_columns = [col for col in required_columns if col not in df_master.columns]
                if missing_columns:
                    messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_columns)}")
                    return

                # Convert Mess Charges to numeric and validate
                df_master['Mess Charges'] = pd.to_numeric(df_master['Mess Charges'], errors='coerce').fillna(0).astype(int)
                invalid_charges = df_master[df_master['Mess Charges'] > 10000]
                if not invalid_charges.empty:
                    messagebox.showerror("Error", f"Invalid Mess Charges found in rows for Employee Codes: {invalid_charges['Employee Code'].tolist()}")
                    return

                new_rows = []
                max_s = self.df_mess['S#'].max() if not self.df_mess.empty else 0
                for idx, row in df_master.iterrows():
                    try:
                        payroll_date = pd.to_datetime(row['PayrollStartDate'], errors='coerce')
                        if pd.isna(payroll_date):
                            messagebox.showerror("Error", f"Invalid PayrollStartDate for employee {row['Employee Code']} at row {idx + 2}. Skipping.")
                            continue
                        salary_month = payroll_date.strftime("%B %Y")
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to parse PayrollStartDate for employee {row['Employee Code']} at row {idx + 2}: {str(e)}. Skipping.")
                        continue

                    code = row['Employee Code']
                    if not ((self.df_mess['Employee Code'] == code) & (self.df_mess['SalaryMonth'] == salary_month)).any():
                        max_s += 1
                        name = row['Employee Name']
                        status = row.get('EmployeeStatus', 'Active')
                        station = row.get('Station', 'PMTF')
                        department = row['DepartmentName']
                        group = row['EmployeeGroup']
                        designation = row['Designation']
                        charges = row['Mess Charges']
                        new_row = {
                            'S#': max_s,
                            'SalaryMonth': salary_month,
                            'Employee Code': code,
                            'Employee Name': name,
                            'Status': status,
                            'StationName': station,
                            'DepartmentName': department,
                            'EmployeeGroup': group,
                            'Designation': designation,
                            'Mess Charges': charges
                        }
                        new_rows.append(new_row)
                        print(f"Added row for {code} for {salary_month} with Mess Charges: {charges}")
                
                if new_rows:
                    new_df = pd.DataFrame(new_rows)
                    new_df['Mess Charges'] = new_df['Mess Charges'].astype(int)
                    self.df_mess = pd.concat([self.df_mess, new_df], ignore_index=True)
                    # Remove duplicates after adding new rows
                    self.df_mess = self.df_mess.drop_duplicates(subset=['Employee Code', 'SalaryMonth'], keep='last')
                    self.save_mess_data()
                    self.month_combo['values'] = sorted(self.df_mess['SalaryMonth'].unique())
                    messagebox.showinfo("Success", f"Added {len(new_rows)} rows for {salary_month}. File saved to {self.mess_file}")
                    print(f"Added {len(new_rows)} rows for {salary_month}")
                else:
                    messagebox.showinfo("Info", "No new entries to add (all employees already exist for the specified month).")
                    print("No new entries added")
                self.display_data()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process the uploaded file: {str(e)}")
        else:
            messagebox.showerror("Error", "No file selected.")

    def display_data(self, df=None):
        if df is None:
            df = self.df_mess
        self.current_df = df
        self.tree.delete(*self.tree.get_children())
        # Debug: Print filtered DataFrame info
        print(f"Displaying data for {len(df)} rows")
        print(f"SalaryMonth values: {df['SalaryMonth'].unique().tolist()}")
        print(f"Mess Charges values: {df['Mess Charges'].tolist()}")
        # Check for invalid or high values
        invalid_charges = df[df['Mess Charges'] > 10000]['Mess Charges'].tolist()
        if invalid_charges:
            print(f"Warning: Abnormally high Mess Charges found: {invalid_charges}")
            messagebox.showwarning("Warning", f"Abnormally high Mess Charges found: {invalid_charges}")
        for idx, row in df.iterrows():
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=tuple(row), tags=(tag,))
        total = df['Mess Charges'].sum() if not df.empty else 0
        print(f"Total Mess Charges calculated: {total}")
        self.total_label.config(text=f"Total Mess Amount: {total}")

    def search_data(self):
        query = self.search_entry.get().strip().lower()
        if not query:
            self.display_data()
            return
        filtered = self.df_mess[
            self.df_mess['Employee Code'].astype(str).str.lower().str.contains(query) |
            self.df_mess['Employee Name'].str.lower().str.contains(query)
        ]
        month = self.month_combo.get()
        if month:
            filtered = filtered[filtered['SalaryMonth'] == month]
        print(f"Search query: {query}, {len(filtered)} rows found")
        self.display_data(filtered)

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Select a row to delete.")
            return
        confirm = messagebox.askyesno("Confirm", "Delete selected rows?")
        if confirm:
            for item in selected:
                values = self.tree.item(item)['values']
                code = values[2]  # Employee Code
                month = values[1]  # SalaryMonth
                self.df_mess = self.df_mess[~((self.df_mess['Employee Code'] == code) & (self.df_mess['SalaryMonth'] == month))]
            self.df_mess['S#'] = range(1, len(self.df_mess) + 1)
            self.save_mess_data()
            self.month_combo['values'] = sorted(self.df_mess['SalaryMonth'].unique())
            self.display_data()

    def edit_selected(self):
        selected = self.tree.selection()
        if not selected or len(selected) != 1:
            messagebox.showwarning("Warning", "Select one row to edit.")
            return
        item = selected[0]
        values = self.tree.item(item)['values']
        edit_win = tk.Toplevel(self)
        edit_win.title("Edit Employee")
        edit_win.configure(bg="#f0f4f8")

        entries = {}
        for i, col in enumerate(self.df_mess.columns):
            ttk.Label(edit_win, text=col).grid(row=i, column=0, sticky="w", padx=10, pady=5)
            entry = ttk.Entry(edit_win, width=50, font=("Helvetica", 10))
            entry.insert(0, str(values[i]) if values[i] is not None else "")
            entry.grid(row=i, column=1, sticky="w", padx=10, pady=5)
            entries[col] = entry

        def save_edit():
            new_values = {col: entries[col].get().strip() for col in self.df_mess.columns}
            original_code = values[2]  # Original Employee Code
            original_month = values[1]  # Original SalaryMonth
            
            mask = (self.df_mess['Employee Code'] == original_code) & (self.df_mess['SalaryMonth'] == original_month)
            if not mask.any():
                messagebox.showerror("Error", f"Entry for {original_code} in {original_month} not found.")
                edit_win.destroy()
                return
            idx = self.df_mess[mask].index[0]
            
            try:
                for col, val in new_values.items():
                    if not val and col not in ['Employee Name', 'DepartmentName', 'Designation']:
                        messagebox.showerror("Error", f"{col} cannot be empty.")
                        return
                    if col == 'Mess Charges':
                        val = int(float(val))  # Convert to int
                        if val > 10000:
                            messagebox.showerror("Error", f"Mess Charges ({val}) cannot be greater than 10000.")
                            return
                    elif col == 'S#':
                        val = int(float(val))
                    else:
                        val = val
                    self.df_mess.at[idx, col] = val
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid value for Mess Charges or S#: {str(e)}")
                return
                
            self.save_mess_data()
            self.display_data()
            self.month_combo['values'] = sorted(self.df_mess['SalaryMonth'].unique())
            messagebox.showinfo("Success", "Employee data updated.")
            edit_win.destroy()

        ttk.Button(edit_win, text="Save", command=save_edit).grid(row=len(self.df_mess.columns), column=1, pady=20)
        ttk.Button(edit_win, text="Delete", command=lambda: [self.delete_selected(), edit_win.destroy()]).grid(row=len(self.df_mess.columns), column=0, pady=20)

    def save_mess_data(self):
        try:
            self.df_mess.to_excel(self.mess_file, index=False)
            print(f"Data saved to {self.mess_file}")
        except PermissionError:
            messagebox.showerror("Error", f"Cannot save to {self.mess_file}. Ensure the file is not open or you have write permissions.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data to {self.mess_file}: {str(e)}")

    def print_data(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showinfo("Info", "No data to print.")
            return
        print_file = os.path.join(self.data_dir, "temp_print_data.xlsx")
        try:
            # Save the current data (filtered or all) to a temporary Excel file
            self.current_df.to_excel(print_file, index=False)
            # Use win32api to send the file to the default printer
            win32api.ShellExecute(0, "print", print_file, None, ".", 0)
            print(f"Sent {print_file} to printer")
            # Delete the temporary file after printing
            try:
                os.remove(print_file)
                print(f"Deleted temporary file: {print_file}")
            except Exception as e:
                print(f"Failed to delete temporary file {print_file}: {str(e)}")
            messagebox.showinfo("Success", "Data sent to printer.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to print data: {str(e)}")