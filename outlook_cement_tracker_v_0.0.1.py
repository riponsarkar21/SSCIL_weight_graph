"""
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
        self.root.title("Cement Delivery Report Tracker")
        self.root.geometry("1400x900")
        self.root.configure(bg="#f0f4f8")
        
        # Initialize database
        self.db_path = "cement_delivery.db"
        self.init_database()
        
        # Email configuration
        self.sender_email = "scale.sscil@sevenringscement.com"
        self.recipient_email = "packing.sscil@sevenringscement.com"
        
        # Variables
        self.current_view = tk.StringVar(value="chart")
        
        self.setup_ui()
        
    def init_database(self):
        """Initialize SQLite database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS delivery_reports (
                date TEXT PRIMARY KEY,
                short INTEGER,
                excess INTEGER,
                per_bag_short_excess REAL,
                email_subject TEXT,
                email_received TEXT,
                UNIQUE(date)
            )
        ''')
        conn.commit()
        conn.close()
        
    def setup_ui(self):
        """Setup the main UI"""
        # Header Frame
        header_frame = tk.Frame(self.root, bg="#2563eb", height=80)
        header_frame.pack(fill=tk.X, pady=0)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame, 
            text="🏭 Cement Delivery Report Tracker",
            font=("Arial", 24, "bold"),
            bg="#2563eb",
            fg="white"
        )
        title_label.pack(pady=20)
        
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
            
            for message in messages:
                try:
                    # Check sender
                    if self.sender_email.lower() not in message.SenderEmailAddress.lower():
                        continue
                    
                    # Check subject
                    subject = message.Subject
                    if "weigh bridge" not in subject.lower() or "report" not in subject.lower():
                        continue
                    
                    # Parse email body
                    body = message.Body
                    parsed_data = self.parse_email_body(body)
                    
                    if parsed_data:
                        report_date = parsed_data['date']
                        received_time = message.ReceivedTime
                        
                        # Keep only the latest email for each date
                        if report_date not in emails_by_date or received_time > emails_by_date[report_date]['received']:
                            emails_by_date[report_date] = {
                                'data': parsed_data,
                                'subject': subject,
                                'received': received_time
                            }
                
                except Exception as e:
                    print(f"Error processing message: {e}")
                    continue
            
            # Insert into database
            for report_date, email_info in emails_by_date.items():
                data = email_info['data']
                try:
                    cursor.execute('''
                        INSERT OR REPLACE INTO delivery_reports 
                        (date, short, excess, per_bag_short_excess, email_subject, email_received)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (
                        data['date'],
                        data['short'],
                        data['excess'],
                        data['per_bag_short_excess'],
                        email_info['subject'],
                        email_info['received'].strftime('%Y-%m-%d %H:%M:%S')
                    ))
                    synced_count += 1
                except Exception as e:
                    print(f"Error inserting record: {e}")
            
            conn.commit()
            conn.close()
            
            return synced_count
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to sync emails: {str(e)}")
            return 0
    
    def parse_email_body(self, body):
        """Parse email body to extract data"""
        try:
            # Extract date (format: DD-MMM-YYYY or DD-Mon-YYYY)
            date_match = re.search(r'Date:\s*(\d{2})-([A-Za-z]{3})-(\d{4})', body)
            if not date_match:
                # Try alternative format in subject
                date_match = re.search(r'(\d{2})\s+([A-Za-z]{3})\s+(\d{4})', body)
            
            if not date_match:
                return None
            
            day, month_str, year = date_match.groups()
            
            # Convert month string to number
            months = {
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            }
            month = months.get(month_str.lower()[:3])
            if not month:
                return None
            
            date_obj = datetime(int(year), month, int(day))
            formatted_date = date_obj.strftime('%Y-%m-%d')
            
            # Find "Daily Report" section
            daily_report_match = re.search(r'Daily Report', body, re.IGNORECASE)
            if not daily_report_match:
                return None
            
            # Get text after "Daily Report"
            daily_section = body[daily_report_match.end():]
            
            # Look for "Monthly to Date Report" to limit our search to Daily section only
            monthly_match = re.search(r'Monthly to Date Report', daily_section, re.IGNORECASE)
            if monthly_match:
                daily_section = daily_section[:monthly_match.start()]
            
            # Now we need to find the row with Short and Excess values
            # The structure is:
            # Total Delivery  Bag Weight  Physical Weight  Short  Excess  Short  Excess
            # 31517          1575850     1577600          320    2070    740    10860
            # But we're in Daily section, so we only have first 5 columns
            
            # Look for the line that has: Short  Excess (column headers)
            # Then find the next line with numbers
            # The pattern after "Physical Weight" column should be: Short Excess
            
            # Find all numbers after the "Short Excess" headers in the table
            # Pattern: Look for "Short" and "Excess" headers, then capture the FIRST TWO numbers that appear
            # This should be in a table row format
            
            # Try to find: Physical Weight followed by Short and Excess columns
            # Match pattern: some numbers (Physical Weight) followed by Short value and Excess value
            table_match = re.search(
                r'Physical\s+Weight\s+Short\s+Excess[\s\S]*?(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)',
                daily_section
            )
            
            if table_match:
                # Groups: 1=Total Delivery, 2=Bag Weight, 3=Physical Weight, 4=Short, 5=Excess
                # We need groups 4 and 5
                short = int(table_match.group(4))
                excess = int(table_match.group(5))
            else:
                # Alternative: Find the data row directly
                # Look for a pattern like: numbers numbers numbers SHORT EXCESS
                # After finding "Short  Excess" header
                header_match = re.search(r'Short\s+Excess', daily_section)
                if header_match:
                    # Get text after header
                    after_header = daily_section[header_match.end():]
                    # Find first line with at least 5 numbers (Total Delivery, Bag Weight, Physical Weight, Short, Excess)
                    row_match = re.search(r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)', after_header)
                    if row_match:
                        short = int(row_match.group(4))
                        excess = int(row_match.group(5))
                    else:
                        return None
                else:
                    return None
            
            # Extract Per Bag Short/Excess from Daily section
            # Should be the FIRST occurrence in daily section
            per_bag_match = re.search(r'Per Bags?\s+Short(?:/Excess)?:\s*(-?\d+\.?\d*)', daily_section, re.IGNORECASE)
            
            if not per_bag_match:
                return None
            
            per_bag = float(per_bag_match.group(1))
            
            return {
                'date': formatted_date,
                'short': short,
                'excess': excess,
                'per_bag_short_excess': per_bag
            }
            
        except Exception as e:
            print(f"Error parsing email: {e}")
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
        
        self.show_short_var = tk.BooleanVar(value=True)
        self.show_excess_var = tk.BooleanVar(value=True)
        self.show_perbag_var = tk.BooleanVar(value=True)
        
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
            SELECT date, short, excess, per_bag_short_excess
            FROM delivery_reports
            WHERE date BETWEEN ? AND ?
            ORDER BY date
        """
        df = pd.read_sql_query(query, conn, params=(from_date, to_date))
        conn.close()
        
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
            ax.plot(df['date'], df['short'], 'o-', color='#ef4444', linewidth=2, markersize=6, label='Daily Short (KG)')
        
        if self.show_excess_var.get():
            ax.plot(df['date'], df['excess'], 's-', color='#10b981', linewidth=2, markersize=6, label='Daily Excess (KG)')
        
        if self.show_perbag_var.get():
            ax2 = ax.twinx()
            ax2.plot(df['date'], df['per_bag_short_excess'], '^-', color='#3b82f6', linewidth=2, markersize=6, label='Per Bag S/E')
            ax2.set_ylabel('Per Bag Short/Excess', fontsize=11, fontweight='bold')
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
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
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
            SELECT date, short, excess, per_bag_short_excess
            FROM delivery_reports
            WHERE date BETWEEN ? AND ?
            ORDER BY date
        """
        df = pd.read_sql_query(query, conn, params=(from_date, to_date))
        conn.close()
        
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
        
        # Create treeview
        tree_frame = tk.Frame(parent_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        vsb = tk.Scrollbar(tree_frame, orient="vertical")
        hsb = tk.Scrollbar(tree_frame, orient="horizontal")
        
        tree = ttk.Treeview(
            tree_frame,
            columns=("Date", "Short", "Excess", "PerBag"),
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
        
        tree.column("Date", width=150, anchor=tk.CENTER)
        tree.column("Short", width=150, anchor=tk.CENTER)
        tree.column("Excess", width=150, anchor=tk.CENTER)
        tree.column("PerBag", width=200, anchor=tk.CENTER)
        
        # Style
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 10), rowheight=30)
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"), background="#4f46e5", foreground="white")
        
        # Insert data
        for idx, row in df.iterrows():
            date_formatted = datetime.strptime(row['date'], '%Y-%m-%d').strftime('%d-%b-%Y')
            tree.insert("", "end", values=(
                date_formatted,
                f"{int(row['short']):,}",
                f"{int(row['excess']):,}",
                f"{row['per_bag_short_excess']:.4f}"
            ), tags=('oddrow' if idx % 2 == 0 else 'evenrow',))
        
        tree.tag_configure('oddrow', background='#f9fafb')
        tree.tag_configure('evenrow', background='white')
        
        # Pack scrollbars and tree
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True)


def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = CementDeliveryTracker(root)
    root.mainloop()


if __name__ == "__main__":
    main()