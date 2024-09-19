import sys
import subprocess
import importlib
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import pyodbc
import threading

def install_package(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def ensure_dependencies():
    required_packages = ['tkinterdnd2', 'pandas', 'openpyxl', 'pyodbc']
    for package in required_packages:
        try:
            importlib.import_module(package)
        except ImportError:
            print(f"Installing {package}...")
            try:
                install_package(package)
            except Exception as e:
                print(f"Failed to install {package}: {str(e)}")
                sys.exit(1)

ensure_dependencies()

class ExcelProcessorGUI:
    def __init__(self, master):
        self.master = master
        master.title("Excel Data Processor Model")
        master.geometry("800x650")
        master.resizable(False, False)
        master.configure(bg='black')

        self.file_path = None
        self.df = None
        self.conn = None
        self.locations = []
        self.create_widgets()

        self.master.drop_target_register(DND_FILES)
        self.master.dnd_bind('<<Drop>>', self.drop)

        self.connect_to_database()
        self.load_locations()

    def create_widgets(self):
        self.frame = tk.Frame(self.master, bg='black')
        self.frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.title_label = tk.Label(self.frame, text="Excel Data Processor Model", fg='yellow', bg='black', font=("Courier", 18, "bold"))
        self.title_label.pack(pady=5)

        self.text_area = tk.Text(self.frame, height=20, width=80, bg='black', fg='yellow', font=("Courier", 10))
        self.text_area.pack(pady=10)
        self.text_area.insert(tk.END, "Drag and drop an Excel file here or click 'Select File'...\n")
        
        # Configure color tags for the text_area
        self.text_area.tag_config('error', foreground='red')
        self.text_area.tag_config('success', foreground='green')
        self.text_area.tag_config('warning', foreground='orange')

        button_frame = tk.Frame(self.frame, bg='black')
        button_frame.pack(pady=5)

        self.select_button = tk.Button(button_frame, text="Select File", command=self.select_file, bg='yellow', fg='black', font=("Courier", 12))
        self.select_button.pack(side=tk.LEFT, padx=5)

        self.read_button = tk.Button(button_frame, text="Read Data", command=self.read_data, bg='yellow', fg='black', font=("Courier", 12))
        self.read_button.pack(side=tk.LEFT, padx=5)

        self.test_button = tk.Button(button_frame, text="Test", command=self.test_data, bg='yellow', fg='black', font=("Courier", 12))
        self.test_button.pack(side=tk.LEFT, padx=5)

        self.process_button = tk.Button(button_frame, text="Process", command=self.process_data, bg='red', fg='white', font=("Courier", 12))
        self.process_button.pack(side=tk.LEFT, padx=5)

        self.location_label = tk.Label(self.frame, text="Location", fg='yellow', bg='black', font=("Courier", 12))
        self.location_label.pack(pady=5)

        self.location_combo = ttk.Combobox(self.frame, state="readonly", font=("Courier", 12))
        self.location_combo.pack(pady=5)
        self.location_combo.set("SELECT LOCATION")

        self.user_label = tk.Label(self.frame, text="User", fg='yellow', bg='black', font=("Courier", 12))
        self.user_label.pack(pady=5)

        self.user_entry = tk.Entry(self.frame, font=("Courier", 12))
        self.user_entry.pack(pady=5)

        self.progress_frame = tk.Frame(self.frame, bg='black')
        self.progress_frame.pack(pady=10, fill=tk.X)

        self.progress_bar = ttk.Progressbar(self.progress_frame, length=700, mode='determinate', style="TProgressbar")
        self.progress_bar.pack(side=tk.LEFT, padx=5)

        self.progress_label = tk.Label(self.progress_frame, text="0%", fg='yellow', bg='black', font=("Courier", 10))
        self.progress_label.pack(side=tk.LEFT, padx=5)

        style = ttk.Style()
        style.theme_use('default')
        style.configure("TProgressbar", thickness=20, troughcolor='black', background='yellow')

    def select_file(self):
        filetypes = (("Excel files", "*.xlsx"), ("All files", "*.*"))
        file = filedialog.askopenfilename(filetypes=filetypes)
        if file:
            self.file_path = file
            self.update_text_area()

    def drop(self, event):
        self.file_path = self.master.tk.splitlist(event.data)[0]
        self.update_text_area()

    def update_text_area(self):
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, f"Selected file: {self.file_path}\n")

    def read_data(self):
        if not self.file_path:
            messagebox.showwarning("Warning", "No file selected.")
            return

        try:
            self.df = pd.read_excel(self.file_path)
            required_columns = ['Location', 'Reference', 'ID', 'Actual_Location']
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            if missing_columns:
                messagebox.showerror("Error", f"Missing columns: {', '.join(missing_columns)}")
                return

            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, self.df.to_string(index=False))
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file: {str(e)}")

    def connect_to_database(self):
        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER=your_server_name;'
            r'DATABASE=your_database_name;'
            r'UID=your_username;'
            r'PWD=your_password;'
        )
        try:
            self.conn = pyodbc.connect(conn_str)
            print("Successfully connected to the database.")
        except pyodbc.Error as e:
            messagebox.showerror("Error", f"Error connecting to database: {str(e)}")
            self.conn = None

    def load_locations(self):
        if self.conn is None:
            return

        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT location FROM locations")
        self.locations = [row.location for row in cursor.fetchall()]
        self.location_combo['values'] = self.locations

    def update_progress(self, current, total, mode='test'):
        progress = int((current / total) * 100)
        self.progress_bar['value'] = progress
        self.progress_label.config(text=f"{progress}%")
        
        style = ttk.Style()
        if mode == 'test':
            style.configure("TProgressbar", background='yellow')
        elif mode == 'process':
            style.configure("TProgressbar", background='red')
        
        self.master.update()

    def test_data(self):
        if self.df is None:
            messagebox.showwarning("Warning", "No data loaded. Please read the Excel file first.")
            return

        if not self.user_entry.get():
            messagebox.showwarning("Warning", "Please enter your username.")
            return

        if self.location_combo.get() == "SELECT LOCATION":
            messagebox.showwarning("Warning", "Please select a valid location.")
            return

        if self.conn is None:
            messagebox.showerror("Error", "No connection to database.")
            return

        threading.Thread(target=self.test_data_thread, daemon=True).start()
    
    def test_data_thread(self):
        results = []
        cursor = self.conn.cursor()
        total_rows = len(self.df)

        for index, row in self.df.iterrows():
            try:
                # Check for blank data
                required_fields = ['ID', 'Reference', 'Actual_Location', 'Location']
                missing_fields = [field for field in required_fields if pd.isna(row[field]) or str(row[field]).strip() == '']
                
                if missing_fields:
                    results.append(('error', f"Row {index + 2}: Missing data for {', '.join(missing_fields)}. Skipping this row."))
                    self.update_progress(index + 1, total_rows, mode='test')
                    continue

                id = str(row['ID'])
                reference = row['Reference']
                actual_location = row['Actual_Location']
                expected_location = row['Location']

                # Query for ID
                cursor.execute("SELECT TOP 1 id, status, reference_id, location_id FROM items WHERE id = ? AND status IS NULL", id)
                item_data = cursor.fetchone()
                
                if item_data is None:
                    results.append(('error', f"Row {index + 2}: ID {id} does not exist or is already processed."))
                    self.update_progress(index + 1, total_rows, mode='test')
                    continue

                # Query for reference
                cursor.execute("SELECT reference_id FROM references WHERE reference = ?", reference)
                ref_data = cursor.fetchone()
                
                if ref_data is None:
                    results.append(('error', f"Row {index + 2}: Reference {reference} does not exist."))
                    self.update_progress(index + 1, total_rows, mode='test')
                    continue
                
                if item_data.reference_id != ref_data.reference_id:
                    results.append(('error', f"Row {index + 2}: ID {id} does not match with reference {reference}."))
                    self.update_progress(index + 1, total_rows, mode='test')
                    continue

                # Query for location
                cursor.execute("SELECT location_id FROM locations WHERE warehouse = ? AND location = ?", 
                            (self.location_combo.get(), actual_location))
                location_data = cursor.fetchone()
                if location_data is None:
                    results.append(('error', f"Row {index + 2}: Location {actual_location} does not exist in warehouse {self.location_combo.get()}."))
                    self.update_progress(index + 1, total_rows, mode='test')
                    continue

                # Check if current location is already correct
                if item_data.location_id == location_data.location_id:
                    results.append(('warning', f"Row {index + 2}: ID {id} is already in the correct location {actual_location}."))
                    self.update_progress(index + 1, total_rows, mode='test')
                    continue

                results.append(('success', f"Row {index + 2}: ID {id} is ready for processing."))

            except Exception as e:
                results.append(('error', f"Error processing row {index + 2}: {str(e)}"))
            
            self.update_progress(index + 1, total_rows, mode='test')

        cursor.close()
        
        self.text_area.delete(1.0, tk.END)
        for tag, message in results:
            self.text_area.insert(tk.END, message + "\n", tag)
                 
    def process_data(self):
        if self.df is None:
            messagebox.showwarning("Warning", "No data loaded. Please read the Excel file first.")
            return

        if not self.user_entry.get():
            messagebox.showwarning("Warning", "Please enter your username.")
            return

        if self.location_combo.get() == "SELECT LOCATION":
            messagebox.showwarning("Warning", "Please select a valid location.")
            return

        if self.conn is None:
            messagebox.showerror("Error", "No connection to database.")
            return

        threading.Thread(target=self.process_data_thread, daemon=True).start()

    def process_data_thread(self):
        results = []
        cursor = self.conn.cursor()

        try:
            warehouse = self.location_combo.get()
            user = self.user_entry.get()
            total_rows = len(self.df)
            
            for index, row in self.df.iterrows():
                try:
                    # Check for blank data
                    required_fields = ['ID', 'Reference', 'Actual_Location', 'Location']
                    missing_fields = [field for field in required_fields if pd.isna(row[field]) or str(row[field]).strip() == '']
                    
                    if missing_fields:
                        results.append(('error', f"Row {index + 2}: Missing data for {', '.join(missing_fields)}. Skipping this row."))
                        self.update_progress(index + 1, total_rows, mode='process')
                        continue

                    id = str(row['ID'])
                    reference = row['Reference']
                    actual_location = row['Actual_Location']
                    expected_location = row['Location']

                    # Get location_id of the actual location
                    cursor.execute("""
                        SELECT location_id 
                        FROM locations 
                        WHERE warehouse = ? AND location = ?
                    """, (warehouse, actual_location))
                    
                    location_data = cursor.fetchone()
                    
                    if location_data is None:
                        results.append(('error', f"Row {index + 2}: Location '{actual_location}' does not exist in warehouse '{warehouse}'."))
                        self.update_progress(index + 1, total_rows, mode='process')
                        continue

                    location_id = location_data[0]

                    # Check if the ID is already in the correct location
                    cursor.execute("""
                        SELECT location_id
                        FROM items
                        WHERE id = ? AND status IS NULL
                    """, id)
                    
                    item_data = cursor.fetchone()
                    
                    if item_data is None:
                        results.append(('error', f"Row {index + 2}: ID {id} does not exist or is already processed."))
                        self.update_progress(index + 1, total_rows, mode='process')
                        continue

                    if item_data.location_id == location_id:
                        results.append(('warning', f"Row {index + 2}: ID {id} is already in the correct location {actual_location}."))
                        self.update_progress(index + 1, total_rows, mode='process')
                        continue

                    # Update item only if not processed and location is different
                    cursor.execute("""
                        UPDATE items 
                        SET location_id = ? 
                        WHERE id = ? AND status IS NULL
                    """, (location_id, id))

                    if cursor.rowcount == 0:
                        results.append(('error', f"Row {index + 2}: ID {id} was not updated. It may not exist or may already be processed."))
                        self.update_progress(index + 1, total_rows, mode='process')
                        continue

                    # Insert movement record
                    cursor.execute("""
                        INSERT INTO movements (movement_type, reference_id, item_id, quantity, units, source_warehouse, 
                        source_location, destination_warehouse, destination_location, date, user, notes, cancelled, 
                        identifier, destination_item_id, identifier_type, terminal_id, sequence, batch, destination_quantity, 
                        movement_block, message_sent, sent_date)
                        SELECT 'RELOCATION', i.reference_id, i.item_id, i.quantity, 
                        CASE WHEN i.quantity > 0 THEN CAST(1 AS FLOAT) ELSE CAST(0 AS FLOAT) END, 
                        ?, ?, ?, ?, GETDATE(), ?, 'location_update_script', 0, NULL, NULL, 
                        NULL, NULL, NULL, i.batch, NULL, NULL, 0, NULL
                        FROM items i 
                        WHERE i.id = ? AND i.status IS NULL
                    """, (warehouse, expected_location, warehouse, actual_location, user, id))

                    if cursor.rowcount > 0:
                        results.append(('success', f"Row {index + 2}: Processed ID {id} successfully."))
                    else:
                        results.append(('error', f"Row {index + 2}: Failed to insert movement for ID {id}. It may not exist or may be processed."))

                except Exception as e:
                    results.append(('error', f"Error processing row {index + 2} (ID {id if 'id' in locals() else 'unknown'}): {str(e)}"))
                    print(f"Detailed error for row {index + 2} in process: {e}")
                    self.conn.rollback()  

                self.update_progress(index + 1, total_rows, mode='process')

            self.conn.commit()
        except Exception as e:
            self.conn.rollback()
            results.append(('error', f"Error processing data: {str(e)}"))
            print(f"Detailed error in process_data_thread: {e}")
        finally:
            cursor.close()
        
        self.text_area.delete(1.0, tk.END)
        for tag, message in results:
            self.text_area.insert(tk.END, message + "\n", tag)
            
def main():
    root = TkinterDnD.Tk()
    ExcelProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
