"""
Last Stable version is v_0.0.25

Cement Delivery Report Tracker - Outlook Integration
Email: packing.sscil@sevenringscement.com
Fetches emails from: scale.sscil@sevenringscement.com

Requirements:
pip install pywin32 pandas matplotlib tkinter pillow
"""

import win32com.client
import re
import sqlite3
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import os

class CementDeliveryTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("Cement Delivery Report Tracker - Ripon Sarkar")
        # Set window to fullscreen dynamically based on screen size
        self.root.attributes('-fullscreen', True)
        self.root.configure(bg="#f0f4f8")
        
        # Bind Escape key to exit fullscreen mode
        self.root.bind("<Escape>", self.toggle_fullscreen)
        
        # Store fullscreen state
        self.fullscreen_state = True
        
        # Schedule close button creation after UI setup
        self.root.after(100, self.create_close_button)
        
        # Initialize database
        self.db_path = "cement_delivery.db"
        self.init_database()
        
        # Email configuration
        self.sender_email = "scale.sscil@sevenringscement.com"
        self.recipient_email = "packing.sscil@sevenringscement.com"
        
        # Variables
        self.current_view = tk.StringVar(value="chart")
        
        # Get version from version.txt
        self.version = self.get_version_from_file()
        
        self.setup_ui()
        
    def init_database(self):
        """Initialize SQLite database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Create table if not exists
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS delivery_reports (
                date TEXT PRIMARY KEY,
                short INTEGER,
                excess INTEGER,
                per_bag_short_excess REAL,
                bag_weight REAL,
                email_subject TEXT,
                email_received TEXT,
                UNIQUE(date)
            )
        ''')
        
        # Check if bag_weight column exists, if not add it
        cursor.execute("PRAGMA table_info(delivery_reports)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'bag_weight' not in columns:
            print("Migrating database: Adding bag_weight column...")
            cursor.execute('ALTER TABLE delivery_reports ADD COLUMN bag_weight REAL')
            
            # Calculate bag_weight for existing records
            cursor.execute('SELECT date, per_bag_short_excess FROM delivery_reports')
            records = cursor.fetchall()
            
            for date, per_bag in records:
                if per_bag is not None:
                    bag_weight = 50.0 - per_bag
                    cursor.execute('UPDATE delivery_reports SET bag_weight = ? WHERE date = ?', (bag_weight, date))
            
            print(f"Migration complete: Updated {len(records)} records with bag_weight")
        
        conn.commit()
        conn.close()
        
    def toggle_fullscreen(self, event=None):
        """Toggle fullscreen mode"""
        self.fullscreen_state = not self.fullscreen_state
        self.root.attributes("-fullscreen", self.fullscreen_state)
        
        # Adjust windowed mode height (reduce by 100px)
        if not self.fullscreen_state:
            # Get current screen dimensions
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            # Set windowed mode with reduced height
            windowed_height = screen_height - 100
            self.root.geometry(f"{screen_width}x{windowed_height}+0+0")
            
            # Hide close button in windowed mode
            if hasattr(self, 'close_button'):
                self.close_button.place_forget()
        else:
            # Show close button in fullscreen mode
            if hasattr(self, 'close_button'):
                self.close_button.place(relx=1.0, rely=0.0, x=-10, y=10, anchor="ne")
        
        return "break"
        
    def create_close_button(self):
        """Create a close button for fullscreen mode"""
        # Only create the button if it doesn't already exist
        if not hasattr(self, 'close_button'):
            self.close_button = tk.Button(
                self.root,
                text="✕",
                command=self.root.destroy,
                bg="#ef4444",
                fg="white",
                font=("Arial", 12, "bold"),
                width=3,
                height=1,
                bd=0,
                highlightthickness=0,
                cursor="hand2"
            )
        
        # Show the close button in fullscreen mode
        if self.fullscreen_state:
            self.close_button.place(relx=1.0, rely=0.0, x=-10, y=10, anchor="ne")
        else:
            self.close_button.place_forget()
        
    def get_version_from_file(self):
        """Extract version from the first line of version.txt"""
        try:
            with open("version.txt", "r") as f:
                first_line = f.readline().strip()
                if first_line:
                    # Extract version from the beginning of the line (before colon or space)
                    version_part = first_line.split(':')[0].strip()
                    if version_part.startswith('v_'):
                        return version_part
                    elif first_line.startswith('v'):
                        return 'v_' + first_line[1:].split(':')[0].strip()
        except Exception as e:
            print(f"Error reading version: {e}")
        return "v_0.1.2"  # fallback version
    
    def setup_ui(self):
        """Setup the main UI"""
        # Header Frame
        header_frame = tk.Frame(self.root, bg="#2563eb", height=80)
        header_frame.pack(fill=tk.X, pady=0)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame, 
            text="🏭 Cement Delivery Report Tracker - by Ripon Sarkar",
            font=("Arial", 24, "bold"),
            bg="#2563eb",
            fg="white"
        )
        title_label.pack(side=tk.LEFT, pady=20, padx=20)
        
        # Version label (clickable)
        version_label = tk.Label(
            header_frame,
            text=self.version,
            font=("Arial", 12, "underline"),
            bg="#2563eb",
            fg="white",
            cursor="hand2"
        )
        version_label.pack(side=tk.RIGHT, pady=20, padx=(20, 60))  # Increased right padding to avoid close button
        version_label.bind("<Button-1>", self.show_version_info)
        
        # Control Panel Frame
        control_frame = tk.Frame(self.root, bg="#e0e7ff", height=100)
        control_frame.pack(fill=tk.X, padx=20, pady=10)
        control_frame.pack_propagate(False)
        
        # Sync Database Button
        sync_btn = tk.Button(
            control_frame,
            text="📥 Sync Database",
            command=self.show_sync_dialog,
            bg="#4f46e5",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        sync_btn.pack(side=tk.LEFT, padx=10, pady=20)
        
        # Rectify Data Button
        rectify_btn = tk.Button(
            control_frame,
            text="✏️ Rectify Data",
            command=self.show_rectify_dialog,
            bg="#f59e0b",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        rectify_btn.pack(side=tk.LEFT, padx=10, pady=20)
        
        # Month Selection
        month_label = tk.Label(
            control_frame,
            text="Select Month:",
            bg="#e0e7ff",
            font=("Arial", 11, "bold")
        )
        month_label.pack(side=tk.LEFT, padx=(30, 5), pady=20)
        
        # Generate month options (last 12 months)
        months = self.generate_months()
        self.month_var = tk.StringVar(value=months[0])
        month_dropdown = ttk.Combobox(
            control_frame,
            textvariable=self.month_var,
            values=months,
            state="readonly",
            width=20,
            font=("Arial", 10)
        )
        month_dropdown.pack(side=tk.LEFT, padx=5, pady=20)
        month_dropdown.bind("<<ComboboxSelected>>", self.on_month_change)
        
        # View Toggle Buttons
        view_frame = tk.Frame(control_frame, bg="#e0e7ff")
        view_frame.pack(side=tk.RIGHT, padx=10, pady=20)
        
        chart_btn = tk.Button(
            view_frame,
            text="📊 Chart View",
            command=lambda: self.switch_view("chart"),
            bg="#4f46e5",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2"
        )
        chart_btn.pack(side=tk.LEFT, padx=5)
        
        table_btn = tk.Button(
            view_frame,
            text="📋 Tabular Form",
            command=lambda: self.switch_view("table"),
            bg="#4f46e5",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2"
        )
        table_btn.pack(side=tk.LEFT, padx=5)
        
        # Content Frame (will hold chart or table)
        self.content_frame = tk.Frame(self.root, bg="white")
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Initialize with chart view
        self.show_chart_view()
        
    def generate_months(self):
        """Generate list of last 12 months"""
        months = []
        today = datetime.now()
        for i in range(12):
            date = today - timedelta(days=30*i)
            month_str = date.strftime("%Y-%m")
            month_label = date.strftime("%B %Y")
            months.append(f"{month_str}|{month_label}")
        return months
    
    def on_month_change(self, event=None):
        """Handle month selection change"""
        if self.current_view.get() == "chart":
            self.show_chart_view()
        else:
            self.show_table_view()
    
    def switch_view(self, view):
        """Switch between chart and table view"""
        self.current_view.set(view)
        if view == "chart":
            self.show_chart_view()
        else:
            self.show_table_view()
    
    def show_sync_dialog(self):
        """Show sync dialog window"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Sync Database")
        dialog.geometry("400x300")
        dialog.configure(bg="white")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")
        
        title = tk.Label(
            dialog,
            text="Sync Database from Outlook",
            font=("Arial", 16, "bold"),
            bg="white"
        )
        title.pack(pady=20)
        
        # From Date
        from_frame = tk.Frame(dialog, bg="white")
        from_frame.pack(pady=10)
        
        tk.Label(from_frame, text="From Date:", bg="white", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=5)
        from_date = tk.Entry(from_frame, width=15, font=("Arial", 10))
        from_date.pack(side=tk.LEFT, padx=5)
        from_date.insert(0, (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"))
        
        tk.Label(from_frame, text="(YYYY-MM-DD)", bg="white", font=("Arial", 8)).pack(side=tk.LEFT)
        
        # To Date
        to_frame = tk.Frame(dialog, bg="white")
        to_frame.pack(pady=10)
        
        tk.Label(to_frame, text="To Date:", bg="white", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=5)
        to_date = tk.Entry(to_frame, width=15, font=("Arial", 10))
        to_date.pack(side=tk.LEFT, padx=5)
        to_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        tk.Label(to_frame, text="(YYYY-MM-DD)", bg="white", font=("Arial", 8)).pack(side=tk.LEFT)
        
        # Progress label
        progress_label = tk.Label(dialog, text="", bg="white", font=("Arial", 10))
        progress_label.pack(pady=20)
        
        # Buttons
        btn_frame = tk.Frame(dialog, bg="white")
        btn_frame.pack(pady=20)
        
        def start_sync():
            from_val = from_date.get()
            to_val = to_date.get()
            
            # Validate dates
            try:
                datetime.strptime(from_val, "%Y-%m-%d")
                datetime.strptime(to_val, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
                return
            
            progress_label.config(text="Syncing... Please wait...")
            dialog.update()
            
            count = self.sync_outlook_emails(from_val, to_val)
            
            progress_label.config(text=f"Successfully synced {count} records!")
            messagebox.showinfo("Success", f"Successfully synced {count} records from Outlook!")
            dialog.destroy()
            
            # Refresh current view
            if self.current_view.get() == "chart":
                self.show_chart_view()
            else:
                self.show_table_view()
        
        sync_btn = tk.Button(
            btn_frame,
            text="Sync Now",
            command=start_sync,
            bg="#4f46e5",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=20,
            pady=8,
            cursor="hand2"
        )
        sync_btn.pack(side=tk.LEFT, padx=10)
        
        cancel_btn = tk.Button(
            btn_frame,
            text="Cancel",
            command=dialog.destroy,
            bg="#6b7280",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=20,
            pady=8,
            cursor="hand2"
        )
        cancel_btn.pack(side=tk.LEFT, padx=10)
    
    def show_rectify_dialog(self):
        """Show rectify data dialog window"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Rectify Data")
        dialog.geometry("900x600")
        dialog.configure(bg="white")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450)
        y = (dialog.winfo_screenheight() // 2) - (300)
        dialog.geometry(f"900x600+{x}+{y}")
        
        title = tk.Label(
            dialog,
            text="Rectify Data - Edit Records",
            font=("Arial", 16, "bold"),
            bg="white"
        )
        title.pack(pady=10)
        
        # Instructions
        instruction = tk.Label(
            dialog,
            text="Double-click on any row to edit. You can modify Short, Excess, and Per Bag values.",
            font=("Arial", 10),
            bg="white",
            fg="#6b7280"
        )
        instruction.pack(pady=5)
        
        # Frame for treeview
        tree_frame = tk.Frame(dialog, bg="white")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Scrollbars
        vsb = tk.Scrollbar(tree_frame, orient="vertical")
        hsb = tk.Scrollbar(tree_frame, orient="horizontal")
        
        # Create treeview
        tree = ttk.Treeview(
            tree_frame,
            columns=("Date", "Short", "Excess", "PerBag", "BagWeight"),
            show="headings",
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            height=15
        )
        
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        
        # Define columns
        tree.heading("Date", text="Date", anchor=tk.CENTER)
        tree.heading("Short", text="Short (KG)", anchor=tk.CENTER)
        tree.heading("Excess", text="Excess (KG)", anchor=tk.CENTER)
        tree.heading("PerBag", text="Per Bag S/E", anchor=tk.CENTER)
        tree.heading("BagWeight", text="Bag Weight (KG)", anchor=tk.CENTER)
        
        tree.column("Date", width=120, anchor=tk.CENTER)
        tree.column("Short", width=150, anchor=tk.CENTER)
        tree.column("Excess", width=150, anchor=tk.CENTER)
        tree.column("PerBag", width=150, anchor=tk.CENTER)
        tree.column("BagWeight", width=150, anchor=tk.CENTER)
        
        # Style
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 10), rowheight=30)
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"))
        
        # Load data
        def load_data():
            # Clear existing items
            for item in tree.get_children():
                tree.delete(item)
            
            # Get all data from database
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                SELECT date, short, excess, per_bag_short_excess, bag_weight
                FROM delivery_reports
                ORDER BY date DESC
            ''')
            records = cursor.fetchall()
            conn.close()
            
            # Insert data
            for idx, record in enumerate(records):
                date, short, excess, per_bag, bag_weight = record
                date_formatted = datetime.strptime(date, '%Y-%m-%d').strftime('%d-%b-%Y')
                
                tree.insert("", "end", values=(
                    date_formatted,
                    f"{int(short):,}",
                    f"{int(excess):,}",
                    f"{per_bag:.4f}",
                    f"{bag_weight:.3f}" if bag_weight else "N/A"
                ), tags=('oddrow' if idx % 2 == 0 else 'evenrow',))
        
        tree.tag_configure('oddrow', background='#f9fafb')
        tree.tag_configure('evenrow', background='white')
        
        # Edit function
        def edit_record(event):
            selected_item = tree.selection()
            if not selected_item:
                return
            
            item = tree.item(selected_item[0])
            values = item['values']
            
            # Parse the date back to YYYY-MM-DD format
            date_str = values[0]  # DD-Mon-YYYY format
            date_obj = datetime.strptime(date_str, '%d-%b-%Y')
            date_formatted = date_obj.strftime('%Y-%m-%d')
            
            # Parse numbers (remove commas)
            short_val = int(str(values[1]).replace(',', ''))
            excess_val = int(str(values[2]).replace(',', ''))
            per_bag_val = float(values[3])
            
            # Create edit dialog
            edit_dialog = tk.Toplevel(dialog)
            edit_dialog.title(f"Edit Record - {date_str}")
            edit_dialog.geometry("400x350")
            edit_dialog.configure(bg="white")
            edit_dialog.transient(dialog)
            edit_dialog.grab_set()
            
            # Center the edit dialog
            edit_dialog.update_idletasks()
            ex = (edit_dialog.winfo_screenwidth() // 2) - 200
            ey = (edit_dialog.winfo_screenheight() // 2) - 175
            edit_dialog.geometry(f"400x350+{ex}+{ey}")
            
            tk.Label(
                edit_dialog,
                text=f"Edit Record for {date_str}",
                font=("Arial", 14, "bold"),
                bg="white"
            ).pack(pady=20)
            
            # Date (read-only)
            date_frame = tk.Frame(edit_dialog, bg="white")
            date_frame.pack(pady=10)
            tk.Label(date_frame, text="Date:", bg="white", font=("Arial", 11, "bold"), width=15, anchor='w').pack(side=tk.LEFT, padx=5)
            tk.Label(date_frame, text=date_str, bg="white", font=("Arial", 11), fg="#6b7280").pack(side=tk.LEFT, padx=5)
            
            # Short
            short_frame = tk.Frame(edit_dialog, bg="white")
            short_frame.pack(pady=10)
            tk.Label(short_frame, text="Short (KG):", bg="white", font=("Arial", 11, "bold"), width=15, anchor='w').pack(side=tk.LEFT, padx=5)
            short_entry = tk.Entry(short_frame, width=15, font=("Arial", 11))
            short_entry.pack(side=tk.LEFT, padx=5)
            short_entry.insert(0, str(short_val))
            
            # Excess
            excess_frame = tk.Frame(edit_dialog, bg="white")
            excess_frame.pack(pady=10)
            tk.Label(excess_frame, text="Excess (KG):", bg="white", font=("Arial", 11, "bold"), width=15, anchor='w').pack(side=tk.LEFT, padx=5)
            excess_entry = tk.Entry(excess_frame, width=15, font=("Arial", 11))
            excess_entry.pack(side=tk.LEFT, padx=5)
            excess_entry.insert(0, str(excess_val))
            
            # Per Bag
            perbag_frame = tk.Frame(edit_dialog, bg="white")
            perbag_frame.pack(pady=10)
            tk.Label(perbag_frame, text="Per Bag S/E:", bg="white", font=("Arial", 11, "bold"), width=15, anchor='w').pack(side=tk.LEFT, padx=5)
            perbag_entry = tk.Entry(perbag_frame, width=15, font=("Arial", 11))
            perbag_entry.pack(side=tk.LEFT, padx=5)
            perbag_entry.insert(0, str(per_bag_val))
            
            # Bag Weight (calculated, read-only display)
            bagweight_frame = tk.Frame(edit_dialog, bg="white")
            bagweight_frame.pack(pady=10)
            tk.Label(bagweight_frame, text="Bag Weight:", bg="white", font=("Arial", 11, "bold"), width=15, anchor='w').pack(side=tk.LEFT, padx=5)
            bagweight_label = tk.Label(bagweight_frame, text=f"{50.0 - per_bag_val:.3f} KG", bg="white", font=("Arial", 11), fg="#7c3aed")
            bagweight_label.pack(side=tk.LEFT, padx=5)
            
            # Update bag weight label when per bag changes
            def update_bagweight_preview(*args):
                try:
                    per_bag = float(perbag_entry.get())
                    bag_wt = 50.0 - per_bag
                    bagweight_label.config(text=f"{bag_wt:.3f} KG")
                except:
                    bagweight_label.config(text="Invalid")
            
            perbag_entry.bind('<KeyRelease>', update_bagweight_preview)
            
            # Save function
            def save_changes():
                try:
                    new_short = int(short_entry.get())
                    new_excess = int(excess_entry.get())
                    new_per_bag = float(perbag_entry.get())
                    new_bag_weight = 50.0 - new_per_bag
                    
                    # Update database
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    cursor.execute('''
                        UPDATE delivery_reports
                        SET short = ?, excess = ?, per_bag_short_excess = ?, bag_weight = ?
                        WHERE date = ?
                    ''', (new_short, new_excess, new_per_bag, new_bag_weight, date_formatted))
                    conn.commit()
                    conn.close()
                    
                    messagebox.showinfo("Success", f"Record for {date_str} updated successfully!")
                    edit_dialog.destroy()
                    load_data()  # Refresh the list
                    
                    # Refresh current view
                    if self.current_view.get() == "chart":
                        self.show_chart_view()
                    else:
                        self.show_table_view()
                    
                except ValueError:
                    messagebox.showerror("Error", "Please enter valid numbers!")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to update record: {str(e)}")
            
            # Delete function
            def delete_record():
                result = messagebox.askyesno(
                    "Confirm Delete",
                    f"Are you sure you want to delete the record for {date_str}?\n\nThis action cannot be undone."
                )
                
                if result:
                    try:
                        conn = sqlite3.connect(self.db_path)
                        cursor = conn.cursor()
                        cursor.execute('DELETE FROM delivery_reports WHERE date = ?', (date_formatted,))
                        conn.commit()
                        conn.close()
                        
                        messagebox.showinfo("Success", f"Record for {date_str} deleted successfully!")
                        edit_dialog.destroy()
                        load_data()  # Refresh the list
                        
                        # Refresh current view
                        if self.current_view.get() == "chart":
                            self.show_chart_view()
                        else:
                            self.show_table_view()
                        
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to delete record: {str(e)}")
            
            # Buttons
            btn_frame = tk.Frame(edit_dialog, bg="white")
            btn_frame.pack(pady=20)
            
            save_btn = tk.Button(
                btn_frame,
                text="💾 Save Changes",
                command=save_changes,
                bg="#10b981",
                fg="white",
                font=("Arial", 10, "bold"),
                padx=15,
                pady=8,
                cursor="hand2"
            )
            save_btn.pack(side=tk.LEFT, padx=5)
            
            delete_btn = tk.Button(
                btn_frame,
                text="🗑️ Delete",
                command=delete_record,
                bg="#ef4444",
                fg="white",
                font=("Arial", 10, "bold"),
                padx=15,
                pady=8,
                cursor="hand2"
            )
            delete_btn.pack(side=tk.LEFT, padx=5)
            
            cancel_btn = tk.Button(
                btn_frame,
                text="Cancel",
                command=edit_dialog.destroy,
                bg="#6b7280",
                fg="white",
                font=("Arial", 10, "bold"),
                padx=15,
                pady=8,
                cursor="hand2"
            )
            cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Bind double-click event
        tree.bind('<Double-Button-1>', edit_record)
        
        # Pack scrollbars and tree
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True)
        
        # Load initial data
        load_data()
        
        # Bottom buttons
        bottom_frame = tk.Frame(dialog, bg="white")
        bottom_frame.pack(pady=10)
        
        refresh_btn = tk.Button(
            bottom_frame,
            text="🔄 Refresh",
            command=load_data,
            bg="#3b82f6",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=8,
            cursor="hand2"
        )
        refresh_btn.pack(side=tk.LEFT, padx=10)
        
        close_btn = tk.Button(
            bottom_frame,
            text="Close",
            command=dialog.destroy,
            bg="#6b7280",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=8,
            cursor="hand2"
        )
        close_btn.pack(side=tk.LEFT, padx=10)
    
    def sync_outlook_emails(self, from_date, to_date):
        """Sync emails from Outlook"""
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
            
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)
            
            # Filter by date
            from_dt = datetime.strptime(from_date, "%Y-%m-%d")
            to_dt = datetime.strptime(to_date, "%Y-%m-%d") + timedelta(days=1)
            
            # Filter messages
            filter_str = f"[ReceivedTime] >= '{from_dt.strftime('%m/%d/%Y')}' AND [ReceivedTime] < '{to_dt.strftime('%m/%d/%Y')}'"
            messages = messages.Restrict(filter_str)
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            synced_count = 0
            emails_by_date = {}
            processed_count = 0
            skipped_count = 0
            
            print(f"\n=== Starting Email Sync ===")
            print(f"Date Range: {from_date} to {to_date}")
            print(f"Total messages in date range: {messages.Count}")
            
            for message in messages:
                try:
                    processed_count += 1
                    
                    # Get sender email - try multiple properties
                    sender_email = ""
                    try:
                        # Try SenderEmailAddress first
                        sender_email = message.SenderEmailAddress
                    except:
                        pass
                    
                    # If SenderEmailAddress doesn't work, try Sender property
                    if not sender_email or "@" not in sender_email:
                        try:
                            sender = message.Sender
                            if sender:
                                sender_email = sender.Address
                        except:
                            pass
                    
                    # If still no email, try SenderName
                    if not sender_email or "@" not in sender_email:
                        try:
                            sender_email = message.SenderName
                        except:
                            pass
                    
                    print(f"\nProcessing email {processed_count}:")
                    print(f"  From: {sender_email}")
                    print(f"  Subject: {message.Subject}")
                    print(f"  Received: {message.ReceivedTime}")
                    
                    # Check sender - be more flexible with matching
                    sender_match = False
                    if sender_email:
                        sender_lower = sender_email.lower()
                        if (self.sender_email.lower() in sender_lower or 
                            "scale.sscil" in sender_lower or
                            "sevenringscement" in sender_lower):
                            sender_match = True
                    
                    if not sender_match:
                        print(f"  ✗ Skipped: Sender doesn't match")
                        skipped_count += 1
                        continue
                    
                    # Check subject - be more flexible
                    subject = message.Subject
                    subject_lower = subject.lower()
                    
                    # Check for various spellings and variations
                    if not (("weigh" in subject_lower and "bridge" in subject_lower) or 
                            "weighbridge" in subject_lower):
                        print(f"  ✗ Skipped: Subject doesn't contain 'weigh bridge'")
                        skipped_count += 1
                        continue
                    
                    if not ("report" in subject_lower or "repot" in subject_lower):
                        print(f"  ✗ Skipped: Subject doesn't contain 'report'")
                        skipped_count += 1
                        continue
                    
                    # Parse email body
                    body = message.Body
                    parsed_data = self.parse_email_body(body)
                    
                    if parsed_data:
                        report_date = parsed_data['date']
                        received_time = message.ReceivedTime
                        
                        print(f"  ✓ Parsed successfully:")
                        print(f"    Date: {report_date}")
                        print(f"    Short: {parsed_data['short']}")
                        print(f"    Excess: {parsed_data['excess']}")
                        print(f"    Per Bag: {parsed_data['per_bag_short_excess']}")
                        
                        # Keep only the latest email for each date
                        if report_date not in emails_by_date or received_time > emails_by_date[report_date]['received']:
                            emails_by_date[report_date] = {
                                'data': parsed_data,
                                'subject': subject,
                                'received': received_time
                            }
                            print(f"    Added to sync list")
                        else:
                            print(f"    Skipped: Earlier email for same date already exists")
                    else:
                        print(f"  ✗ Failed to parse email body")
                        skipped_count += 1
                
                except Exception as e:
                    print(f"  ✗ Error processing message: {e}")
                    import traceback
                    traceback.print_exc()
                    skipped_count += 1
                    continue
            
            print(f"\n=== Inserting into Database ===")
            # Insert into database
            for report_date, email_info in emails_by_date.items():
                data = email_info['data']
                try:
                    cursor.execute('''
                        INSERT OR REPLACE INTO delivery_reports 
                        (date, short, excess, per_bag_short_excess, bag_weight, email_subject, email_received)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        data['date'],
                        data['short'],
                        data['excess'],
                        data['per_bag_short_excess'],
                        data['bag_weight'],
                        email_info['subject'],
                        email_info['received'].strftime('%Y-%m-%d %H:%M:%S')
                    ))
                    synced_count += 1
                    print(f"  ✓ Inserted: {data['date']}")
                except Exception as e:
                    print(f"  ✗ Error inserting {data['date']}: {e}")
            
            conn.commit()
            conn.close()
            
            print(f"\n=== Sync Complete ===")
            print(f"Total processed: {processed_count}")
            print(f"Skipped: {skipped_count}")
            print(f"Successfully synced: {synced_count}")
            
            return synced_count
            
        except Exception as e:
            print(f"\nError during sync: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to sync emails: {str(e)}")
            return 0
    
    def parse_email_body(self, body):
        """Parse email body to extract data"""
        try:
            print(f"    Parsing email body...")
            
            # Extract date (format: DD-MMM-YYYY or DD/MMM/YYYY or DD-Mon-YYYY or DD/Mon/YYYY)
            date_match = re.search(r'Date:\s*(\d{2})[-/]([A-Za-z]{3})[-/](\d{4})', body)
            if not date_match:
                # Try alternative format in subject
                date_match = re.search(r'(\d{2})[-/\s]+([A-Za-z]{3})[-/\s]+(\d{4})', body)
            
            if not date_match:
                print(f"    ✗ Failed: Could not find date in format DD-MMM-YYYY")
                return None
            
            day, month_str, year = date_match.groups()
            print(f"    Found date: {day}-{month_str}-{year}")
            
            # Convert month string to number
            months = {
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            }
            month = months.get(month_str.lower()[:3])
            if not month:
                print(f"    ✗ Failed: Invalid month name '{month_str}'")
                return None
            
            date_obj = datetime(int(year), month, int(day))
            formatted_date = date_obj.strftime('%Y-%m-%d')
            print(f"    Formatted date: {formatted_date}")
            
            # The email has a two-column layout:
            # Daily Report | Monthly to Date Report
            # So we need to find the "Delivery Information: Bag Cement" section
            # and extract the FIRST set of Short/Excess values (which is Daily)
            
            # Find "Delivery Information: Bag Cement"
            delivery_info_match = re.search(r'Delivery Information:\s*Bag Cement', body, re.IGNORECASE)
            if not delivery_info_match:
                print(f"    ✗ Failed: Could not find 'Delivery Information: Bag Cement'")
                return None
            
            print(f"    Found 'Delivery Information: Bag Cement' section")
            
            # Get text after this section
            after_delivery_info = body[delivery_info_match.end():]
            
            # Look for the header row: "Total Delivery  Bag Weight  Physical Weight  Short  Excess"
            # This appears twice (once for Daily, once for Monthly)
            header_pattern = r'Total\s+Delivery\s+Bag\s+Weight\s+Physical\s+Weight\s+Short\s+Excess'
            header_matches = list(re.finditer(header_pattern, after_delivery_info, re.IGNORECASE))
            
            if len(header_matches) == 0:
                print(f"    ✗ Failed: Could not find table header")
                return None
            
            print(f"    Found {len(header_matches)} table header(s)")
            
            # Get the FIRST data row after the FIRST header (this is Daily Report)
            first_header = header_matches[0]
            after_first_header = after_delivery_info[first_header.end():]
            
            # The data row has 5 or 7 numbers (depending on if Monthly columns are included)
            # Format: Total_Delivery Bag_Weight Physical_Weight Short Excess [Short_Monthly Excess_Monthly]
            # We want the first 5 numbers
            
            # Look for a line with numbers - could be on same line or next lines
            # Try to find: 5 consecutive numbers (allowing for whitespace and newlines)
            data_match = re.search(
                r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)',
                after_first_header
            )
            
            if not data_match:
                print(f"    ✗ Failed: Could not find data row with 5 numbers")
                print(f"    After header text (first 500 chars): {after_first_header[:500]}")
                return None
            
            # Extract values - the 4th and 5th numbers are Short and Excess
            total_delivery = int(data_match.group(1))
            bag_weight = int(data_match.group(2))
            physical_weight = int(data_match.group(3))
            short = int(data_match.group(4))
            excess = int(data_match.group(5))
            
            print(f"    Found data row:")
            print(f"      Total Delivery: {total_delivery}")
            print(f"      Bag Weight: {bag_weight}")
            print(f"      Physical Weight: {physical_weight}")
            print(f"      Short: {short}")
            print(f"      Excess: {excess}")
            
            # Extract Per Bag Short/Excess
            # Look for the FIRST occurrence (Daily Report value)
            per_bag_match = re.search(r'Per Bags?\s+Short(?:/Excess)?:\s*(-?\d+\.?\d*)', body, re.IGNORECASE)
            
            if not per_bag_match:
                print(f"    ✗ Failed: Could not find 'Per Bag Short/Excess' value")
                return None
            
            per_bag = float(per_bag_match.group(1))
            print(f"    Found Per Bag Short/Excess: {per_bag}")
            
            # Calculate Bag Weight = 50 - (Per Bag Short/Excess)
            bag_weight = 50.0 - per_bag
            print(f"    Calculated Bag Weight: {bag_weight} KG")
            
            return {
                'date': formatted_date,
                'short': short,
                'excess': excess,
                'per_bag_short_excess': per_bag,
                'bag_weight': bag_weight
            }
            
        except Exception as e:
            print(f"    ✗ Exception during parsing: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def get_date_range_from_month(self, month_str):
        """Get date range from selected month"""
        year_month = month_str.split('|')[0]
        year, month = map(int, year_month.split('-'))
        
        from_date = f"{year}-{month:02d}-01"
        
        # Last day of month
        if month == 12:
            next_month = datetime(year + 1, 1, 1)
        else:
            next_month = datetime(year, month + 1, 1)
        
        last_day = (next_month - timedelta(days=1)).day
        to_date = f"{year}-{month:02d}-{last_day:02d}"
        
        return from_date, to_date
    
    def show_chart_view(self):
        """Display chart view"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Date range selector frame
        date_frame = tk.Frame(self.content_frame, bg="#f3f4f6", height=60)
        date_frame.pack(fill=tk.X, padx=10, pady=10)
        date_frame.pack_propagate(False)
        
        tk.Label(date_frame, text="📅", bg="#f3f4f6", font=("Arial", 14)).pack(side=tk.LEFT, padx=10)
        
        # Get default dates from selected month
        from_date, to_date = self.get_date_range_from_month(self.month_var.get())
        
        tk.Label(date_frame, text="From:", bg="#f3f4f6", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        from_entry = tk.Entry(date_frame, width=12, font=("Arial", 10))
        from_entry.pack(side=tk.LEFT, padx=5)
        from_entry.insert(0, from_date)
        
        tk.Label(date_frame, text="To:", bg="#f3f4f6", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        to_entry = tk.Entry(date_frame, width=12, font=("Arial", 10))
        to_entry.pack(side=tk.LEFT, padx=5)
        to_entry.insert(0, to_date)
        
        # Checkboxes frame
        check_frame = tk.Frame(self.content_frame, bg="white")
        check_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.show_short_var = tk.BooleanVar(value=False)
        self.show_excess_var = tk.BooleanVar(value=False)
        self.show_perbag_var = tk.BooleanVar(value=False)
        self.show_bagweight_var = tk.BooleanVar(value=True)
        
        def update_chart():
            self.plot_chart(canvas_frame, from_entry.get(), to_entry.get())
        
        tk.Checkbutton(
            check_frame, 
            text="Daily Short (KG)", 
            variable=self.show_short_var,
            command=update_chart,
            bg="white",
            font=("Arial", 10),
            fg="#ef4444"
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Checkbutton(
            check_frame, 
            text="Daily Excess (KG)", 
            variable=self.show_excess_var,
            command=update_chart,
            bg="white",
            font=("Arial", 10),
            fg="#10b981"
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Checkbutton(
            check_frame, 
            text="Per Bag Short/Excess", 
            variable=self.show_perbag_var,
            command=update_chart,
            bg="white",
            font=("Arial", 10),
            fg="#3b82f6"
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Checkbutton(
            check_frame, 
            text="Bag Weight (KG)", 
            variable=self.show_bagweight_var,
            command=update_chart,
            bg="white",
            font=("Arial", 10),
            fg="#8b5cf6"
        ).pack(side=tk.LEFT, padx=10)
        
        update_btn = tk.Button(
            check_frame,
            text="Update Chart",
            command=update_chart,
            bg="#4f46e5",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=5,
            cursor="hand2"
        )
        update_btn.pack(side=tk.RIGHT, padx=10)
        
        # Chart canvas frame
        canvas_frame = tk.Frame(self.content_frame, bg="white")
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Initial plot
        self.plot_chart(canvas_frame, from_date, to_date)
    
    def plot_chart(self, parent_frame, from_date, to_date):
        """Plot the chart"""
        # Clear previous chart
        for widget in parent_frame.winfo_children():
            widget.destroy()
        
        # Get data
        conn = sqlite3.connect(self.db_path)
        query = """
            SELECT date, short, excess, per_bag_short_excess, bag_weight
            FROM delivery_reports
            WHERE date BETWEEN ? AND ?
            ORDER BY date
        """
        df = pd.read_sql_query(query, conn, params=(from_date, to_date))
        conn.close()
        
        # For old records without bag_weight, calculate it
        if 'bag_weight' in df.columns:
            df['bag_weight'] = df.apply(
                lambda row: row['bag_weight'] if pd.notna(row['bag_weight']) else 50.0 - row['per_bag_short_excess'],
                axis=1
            )
        
        if df.empty:
            no_data_label = tk.Label(
                parent_frame,
                text="📊 No data available for the selected date range\n\nClick 'Sync Database' to load data",
                font=("Arial", 14),
                bg="white",
                fg="#6b7280"
            )
            no_data_label.pack(expand=True)
            return
        
        # Create figure
        fig = Figure(figsize=(12, 6), dpi=100)
        ax = fig.add_subplot(111)
        
        df['date'] = pd.to_datetime(df['date'])
        
        # Plot lines based on checkboxes
        if self.show_short_var.get():
            ax.plot(df['date'], df['short'], 'o-', color='#ef4444', linewidth=2, markersize=8, label='Daily Short (KG)')
        
        if self.show_excess_var.get():
            ax.plot(df['date'], df['excess'], 's-', color='#10b981', linewidth=2, markersize=8, label='Daily Excess (KG)')
        
        # Use secondary Y-axis for Per Bag and Bag Weight
        ax2 = None
        if self.show_perbag_var.get() or self.show_bagweight_var.get():
            ax2 = ax.twinx()
            
            if self.show_perbag_var.get():
                ax2.plot(df['date'], df['per_bag_short_excess'], '^-', color='#3b82f6', linewidth=2, markersize=8, label='Per Bag S/E')
            
            if self.show_bagweight_var.get():
                ax2.plot(df['date'], df['bag_weight'], 'd-', color='#8b5cf6', linewidth=2, markersize=8, label='Bag Weight (KG)')
            
            ax2.set_ylabel('Per Bag S/E / Bag Weight (KG)', fontsize=11, fontweight='bold')
            ax2.legend(loc='upper right')
            ax2.grid(False)
        
        ax.set_xlabel('Date', fontsize=11, fontweight='bold')
        ax.set_ylabel('Weight (KG)', fontsize=11, fontweight='bold')
        ax.set_title('Cement Delivery Report - Daily Analysis', fontsize=14, fontweight='bold')
        ax.legend(loc='upper left')
        ax.grid(True, alpha=0.3)
        
        fig.autofmt_xdate()
        fig.tight_layout()
        
        # Embed chart in tkinter
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.draw()
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)
        
        # Create tooltip label (initially hidden)
        tooltip = tk.Label(
            parent_frame,
            text="",
            bg='#1f2937',
            fg='white',
            font=('Arial', 9),
            relief=tk.RAISED,
            borderwidth=2,
            justify=tk.LEFT,
            padx=10,
            pady=8
        )
        
        # Store tooltip reference
        tooltip.place_forget()
        
        def find_nearest_point(event):
            """Find the nearest data point to the mouse cursor"""
            if event.inaxes not in [ax, ax2]:
                return None
            
            # Get mouse position in data coordinates
            mouse_x = event.xdata
            mouse_y = event.ydata
            
            if mouse_x is None or mouse_y is None:
                return None
            
            # Convert matplotlib date to datetime
            try:
                mouse_date = pd.to_datetime(mouse_x, unit='D', origin='unix')
            except:
                return None
            
            # Find nearest date in dataframe
            date_diffs = abs(df['date'] - mouse_date)
            nearest_idx = date_diffs.idxmin()
            
            # Check if mouse is close enough to this point (within reasonable range)
            # Get the point's x position in display coordinates
            point_date = df.iloc[nearest_idx]['date']
            
            # Convert to matplotlib format for comparison
            point_x = pd.Timestamp(point_date).value / 10**9 / 86400 + 719163
            
            # Check if mouse is within reasonable x-range (e.g., 5% of data range)
            x_range = max(df['date']).value - min(df['date']).value
            x_threshold = x_range * 0.05 / 10**9 / 86400
            
            if abs(mouse_x - point_x) > x_threshold:
                return None
            
            return nearest_idx
        
        def on_hover(event):
            """Handle mouse hover event"""
            nearest_idx = find_nearest_point(event)
            
            if nearest_idx is not None:
                # Get data for this point
                date_val = df.iloc[nearest_idx]['date']
                short_val = df.iloc[nearest_idx]['short']
                excess_val = df.iloc[nearest_idx]['excess']
                perbag_val = df.iloc[nearest_idx]['per_bag_short_excess']
                bagweight_val = df.iloc[nearest_idx]['bag_weight']
                
                # Format date
                date_str = date_val.strftime('%d-%b-%Y')
                
                # Create tooltip text with colors (using Unicode characters for colored circles)
                tooltip_text = f"📅 {date_str}\n"
                tooltip_text += "─" * 30 + "\n"
                tooltip_text += f"🔴 Short:           {int(short_val):,} KG\n"
                tooltip_text += f"🟢 Excess:         {int(excess_val):,} KG\n"
                tooltip_text += f"🔵 Per Bag S/E:  {perbag_val:.4f}\n"
                tooltip_text += f"🟣 Bag Weight:   {bagweight_val:.3f} KG"
                
                tooltip.config(text=tooltip_text)
                
                # Position tooltip near cursor
                x = event.x + 15
                y = event.y + 15
                
                # Make sure tooltip doesn't go off screen
                tooltip.update_idletasks()
                tooltip_width = tooltip.winfo_reqwidth()
                tooltip_height = tooltip.winfo_reqheight()
                
                canvas_width = canvas_widget.winfo_width()
                canvas_height = canvas_widget.winfo_height()
                
                if x + tooltip_width > canvas_width:
                    x = event.x - tooltip_width - 15
                
                if y + tooltip_height > canvas_height:
                    y = event.y - tooltip_height - 15
                
                tooltip.place(x=x, y=y)
            else:
                tooltip.place_forget()
        
        def on_leave(event):
            """Hide tooltip when mouse leaves chart"""
            tooltip.place_forget()
        
        # Connect hover events
        canvas.mpl_connect('motion_notify_event', on_hover)
        canvas.mpl_connect('axes_leave_event', on_leave)
        canvas.mpl_connect('figure_leave_event', on_leave)
    
    def show_table_view(self):
        """Display table view"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Date range selector frame
        date_frame = tk.Frame(self.content_frame, bg="#f3f4f6", height=60)
        date_frame.pack(fill=tk.X, padx=10, pady=10)
        date_frame.pack_propagate(False)
        
        tk.Label(date_frame, text="📅", bg="#f3f4f6", font=("Arial", 14)).pack(side=tk.LEFT, padx=10)
        
        # Get default dates from selected month
        from_date, to_date = self.get_date_range_from_month(self.month_var.get())
        
        tk.Label(date_frame, text="From:", bg="#f3f4f6", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        from_entry = tk.Entry(date_frame, width=12, font=("Arial", 10))
        from_entry.pack(side=tk.LEFT, padx=5)
        from_entry.insert(0, from_date)
        
        tk.Label(date_frame, text="To:", bg="#f3f4f6", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        to_entry = tk.Entry(date_frame, width=12, font=("Arial", 10))
        to_entry.pack(side=tk.LEFT, padx=5)
        to_entry.insert(0, to_date)
        
        def update_table():
            self.display_table(table_container, from_entry.get(), to_entry.get(), summary_frame)
        
        update_btn = tk.Button(
            date_frame,
            text="Update Table",
            command=update_table,
            bg="#4f46e5",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=5,
            cursor="hand2"
        )
        update_btn.pack(side=tk.RIGHT, padx=10)
        
        # Summary frame
        summary_frame = tk.Frame(self.content_frame, bg="white")
        summary_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Table container
        table_container = tk.Frame(self.content_frame, bg="white")
        table_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Initial display
        self.display_table(table_container, from_date, to_date, summary_frame)
    
    def display_table(self, parent_frame, from_date, to_date, summary_frame):
        """Display data in table format"""
        # Clear previous widgets
        for widget in parent_frame.winfo_children():
            widget.destroy()
        for widget in summary_frame.winfo_children():
            widget.destroy()
        
        # Get data
        conn = sqlite3.connect(self.db_path)
        query = """
            SELECT date, short, excess, per_bag_short_excess, bag_weight
            FROM delivery_reports
            WHERE date BETWEEN ? AND ?
            ORDER BY date
        """
        df = pd.read_sql_query(query, conn, params=(from_date, to_date))
        conn.close()
        
        # For old records without bag_weight, calculate it
        if 'bag_weight' in df.columns:
            df['bag_weight'] = df.apply(
                lambda row: row['bag_weight'] if pd.notna(row['bag_weight']) else 50.0 - row['per_bag_short_excess'],
                axis=1
            )
        
        if df.empty:
            no_data_label = tk.Label(
                parent_frame,
                text="📋 No data available for the selected date range\n\nClick 'Sync Database' to load data",
                font=("Arial", 14),
                bg="white",
                fg="#6b7280"
            )
            no_data_label.pack(expand=True)
            return
        
        # Calculate totals
        total_short = df['short'].sum()
        total_excess = df['excess'].sum()
        avg_per_bag = df['per_bag_short_excess'].mean()
        avg_bag_weight = df['bag_weight'].mean()
        
        # Display summary cards
        card1 = tk.Frame(summary_frame, bg="#fee2e2", relief=tk.RAISED, borderwidth=2)
        card1.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        tk.Label(card1, text="Total Short", bg="#fee2e2", font=("Arial", 10)).pack(pady=5)
        tk.Label(card1, text=f"{total_short:,} KG", bg="#fee2e2", font=("Arial", 16, "bold"), fg="#dc2626").pack(pady=5)
        
        card2 = tk.Frame(summary_frame, bg="#d1fae5", relief=tk.RAISED, borderwidth=2)
        card2.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        tk.Label(card2, text="Total Excess", bg="#d1fae5", font=("Arial", 10)).pack(pady=5)
        tk.Label(card2, text=f"{total_excess:,} KG", bg="#d1fae5", font=("Arial", 16, "bold"), fg="#059669").pack(pady=5)
        
        card3 = tk.Frame(summary_frame, bg="#dbeafe", relief=tk.RAISED, borderwidth=2)
        card3.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        tk.Label(card3, text="Avg Per Bag S/E", bg="#dbeafe", font=("Arial", 10)).pack(pady=5)
        tk.Label(card3, text=f"{avg_per_bag:.4f}", bg="#dbeafe", font=("Arial", 16, "bold"), fg="#2563eb").pack(pady=5)
        
        card4 = tk.Frame(summary_frame, bg="#e9d5ff", relief=tk.RAISED, borderwidth=2)
        card4.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        tk.Label(card4, text="Avg Bag Weight", bg="#e9d5ff", font=("Arial", 10)).pack(pady=5)
        tk.Label(card4, text=f"{avg_bag_weight:.3f} KG", bg="#e9d5ff", font=("Arial", 16, "bold"), fg="#7c3aed").pack(pady=5)
        
        # Create treeview
        tree_frame = tk.Frame(parent_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        vsb = tk.Scrollbar(tree_frame, orient="vertical")
        hsb = tk.Scrollbar(tree_frame, orient="horizontal")
        
        tree = ttk.Treeview(
            tree_frame,
            columns=("Date", "Short", "Excess", "PerBag", "BagWeight"),
            show="headings",
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            height=20
        )
        
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        
        # Define columns
        tree.heading("Date", text="Date", anchor=tk.CENTER)
        tree.heading("Short", text="Short (KG)", anchor=tk.CENTER)
        tree.heading("Excess", text="Excess (KG)", anchor=tk.CENTER)
        tree.heading("PerBag", text="Per Bag Short/Excess", anchor=tk.CENTER)
        tree.heading("BagWeight", text="Bag Weight (KG)", anchor=tk.CENTER)
        
        tree.column("Date", width=120, anchor=tk.CENTER)
        tree.column("Short", width=120, anchor=tk.CENTER)
        tree.column("Excess", width=120, anchor=tk.CENTER)
        tree.column("PerBag", width=150, anchor=tk.CENTER)
        tree.column("BagWeight", width=150, anchor=tk.CENTER)
        
        # Style
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 10), rowheight=30)
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"), background="#4f46e5", foreground="Blue")
        
        # Insert data
        for idx, row in df.iterrows():
            date_formatted = datetime.strptime(row['date'], '%Y-%m-%d').strftime('%d-%b-%Y')
            tree.insert("", "end", values=(
                date_formatted,
                f"{int(row['short']):,}",
                f"{int(row['excess']):,}",
                f"{row['per_bag_short_excess']:.4f}",
                f"{row['bag_weight']:.3f}"
            ), tags=('oddrow' if idx % 2 == 0 else 'evenrow',))
        
        tree.tag_configure('oddrow', background='#f9fafb')
        tree.tag_configure('evenrow', background='white')
        
        # Pack scrollbars and tree
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True)
    
    def show_version_info(self, event=None):
        """Show version information in a popup window"""
        try:
            with open("version.txt", "r") as f:
                version_content = f.read()
        except FileNotFoundError:
            version_content = "Version information not available."
        except Exception as e:
            version_content = f"Error reading version information: {str(e)}"
        
        # Create popup window
        popup = tk.Toplevel(self.root)
        popup.title("Version Information")
        popup.geometry("500x400")
        popup.configure(bg="white")
        popup.transient(self.root)
        popup.grab_set()
        
        # Center the popup
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() // 2) - (500 // 2)
        y = (popup.winfo_screenheight() // 2) - (400 // 2)
        popup.geometry(f"500x400+{x}+{y}")
        
        # Title
        title = tk.Label(
            popup,
            text="Version Information",
            font=("Arial", 16, "bold"),
            bg="white",
            fg="#2563eb"
        )
        title.pack(pady=20)
        
        # Text widget for version content
        text_frame = tk.Frame(popup, bg="white")
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Arial", 10),
            bg="#f8fafc",
            fg="#1e293b",
            padx=10,
            pady=10,
            spacing1=2,
            spacing3=2
        )
        text_widget.insert(tk.END, version_content)
        text_widget.config(state=tk.DISABLED)
        
        scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Close button
        close_btn = tk.Button(
            popup,
            text="Close",
            command=popup.destroy,
            bg="#6b7280",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=8,
            cursor="hand2"
        )
        close_btn.pack(pady=20)


def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = CementDeliveryTracker(root)
    root.mainloop()


if __name__ == "__main__":
    main()