"""
Dashboard for Cement Delivery Tracking Applications
Provides buttons to run various versions of the tracking software
"""

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os
from threading import Thread

class ApplicationDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Cement Delivery Tracker Dashboard")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f4f8")
        
        # Store process references and status labels
        self.processes = {}
        self.status_labels = {}
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the dashboard UI"""
        # Header
        header_frame = tk.Frame(self.root, bg="#2563eb", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="🏭 Cement Delivery Tracker Dashboard",
            font=("Arial", 20, "bold"),
            bg="#2563eb",
            fg="white"
        )
        title_label.pack(pady=20)
        
        # Description
        desc_label = tk.Label(
            self.root,
            text="Click on any button below to run the corresponding application",
            font=("Arial", 12),
            bg="#f0f4f8",
            fg="#334155"
        )
        desc_label.pack(pady=10)
        
        # Main content frame with scrollable canvas
        canvas_frame = tk.Frame(self.root, bg="#f0f4f8")
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Canvas and scrollbar
        self.canvas = tk.Canvas(canvas_frame, bg="#f0f4f8", highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#f0f4f8")
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas scrolling
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)
        
        # Create buttons for each application
        self.create_app_buttons()
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg="#cbd5e1",
            fg="#1e293b"
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def create_app_buttons(self):
        """Create buttons for each application"""
        # Define applications with their display names and file paths
        self.applications = [
            {
                "name": "Export DB to CSV",
                "file": "export_db_to_csv.py",
                "desc": "Exports database records to CSV file"
            },
            {
                "name": "Outlook Tracker v0.0.1",
                "file": "outlook_cement_tracker_v_0.0.1.py",
                "desc": "Original version of the Outlook tracker"
            },
            {
                "name": "Outlook Tracker v0.0.2",
                "file": "outlook_cement_tracker_v_0.0.2.py",
                "desc": "Updated version with improvements"
            },
            {
                "name": "Outlook Tracker v0.0.6",
                "file": "outlook_cement_tracker_v_0.0.6.py",
                "desc": "Enhanced features version"
            },
            {
                "name": "Outlook Tracker v0.0.7",
                "file": "outlook_cement_tracker_v_0.0.7.py",
                "desc": "Database optimization version"
            },
            {
                "name": "Outlook Tracker v0.0.14",
                "file": "outlook_cement_tracker_v_0.0.14.py",
                "desc": "UI enhancements version"
            },
            {
                "name": "Outlook Tracker v0.0.20",
                "file": "outlook_cement_tracker_v_0.0.20.py",
                "desc": "Performance improvements version"
            },
            {
                "name": "Outlook Tracker v0.0.22",
                "file": "outlook_cement_tracker_v_0.0.22.py",
                "desc": "Advanced features version"
            },
            {
                "name": "Outlook Tracker v0.0.23 XX",
                "file": "outlook_cement_tracker_v_0.0.23_xx.py",
                "desc": "Experimental version"
            },
            {
                "name": "Outlook Tracker v0.0.24",
                "file": "outlook_cement_tracker_v_0.0.24.py",
                "desc": "Latest stable version"
            },
            {
                "name": "Outlook Tracker (Main)",
                "file": "outlook_cement_tracker.py",
                "desc": "Current main version"
            }
        ]
        
        # Create buttons in a grid layout
        row, col = 0, 0
        max_cols = 2
        
        for i, app in enumerate(self.applications):
            # Create a frame for each application
            app_frame = tk.Frame(
                self.scrollable_frame,
                bg="white",
                relief=tk.RAISED,
                borderwidth=1
            )
            app_frame.grid(row=row, column=col, padx=10, pady=10, sticky="ew")
            app_frame.columnconfigure(1, weight=1)
            
            # Application name
            name_label = tk.Label(
                app_frame,
                text=app["name"],
                font=("Arial", 12, "bold"),
                bg="white",
                fg="#1e40af"
            )
            name_label.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 0))
            
            # Description
            desc_label = tk.Label(
                app_frame,
                text=app["desc"],
                font=("Arial", 9),
                bg="white",
                fg="#64748b",
                wraplength=300,
                justify=tk.LEFT
            )
            desc_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 10))
            
            # Button frame
            btn_frame = tk.Frame(app_frame, bg="white")
            btn_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")
            
            # Run button
            run_btn = tk.Button(
                btn_frame,
                text="▶ Run",
                command=lambda f=app["file"]: self.run_application(f),
                bg="#10b981",
                fg="white",
                font=("Arial", 10, "bold"),
                padx=15,
                cursor="hand2"
            )
            run_btn.pack(side=tk.LEFT, padx=(0, 5))
            
            # Stop button
            stop_btn = tk.Button(
                btn_frame,
                text="⏹ Stop",
                command=lambda f=app["file"]: self.stop_application(f),
                bg="#ef4444",
                fg="white",
                font=("Arial", 10, "bold"),
                padx=15,
                cursor="hand2"
            )
            stop_btn.pack(side=tk.LEFT, padx=(5, 0))
            
            # Status indicator
            status_label = tk.Label(
                btn_frame,
                text="●",
                font=("Arial", 14),
                fg="#94a3b8",
                bg="white"
            )
            status_label.pack(side=tk.LEFT, padx=(10, 0))
            
            # Store reference to status label
            self.status_labels[app["file"]] = status_label
            
            # Update column and row for grid layout
            col += 1
            if col >= max_cols:
                col = 0
                row += 1
                
        # Configure column weights for responsive design
        self.scrollable_frame.columnconfigure(0, weight=1)
        self.scrollable_frame.columnconfigure(1, weight=1)
        
    def run_application(self, filename):
        """Run the specified Python application"""
        if not os.path.exists(filename):
            messagebox.showerror("Error", f"File {filename} not found!")
            return
            
        # Check if already running
        if filename in self.processes and self.processes[filename].poll() is None:
            messagebox.showinfo("Info", f"{filename} is already running!")
            return
            
        try:
            # Update status
            self.update_app_status(filename, "running")
            self.status_var.set(f"Running {filename}...")
            
            # Run the application in a separate thread to avoid blocking UI
            thread = Thread(target=self._run_process, args=(filename,))
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run {filename}: {str(e)}")
            self.update_app_status(filename, "error")
            self.status_var.set("Ready")
            
    def _run_process(self, filename):
        """Internal method to run the process"""
        try:
            # Run the Python file
            self.processes[filename] = subprocess.Popen([sys.executable, filename])
            # Wait for completion
            self.processes[filename].wait()
            
            # Update status based on exit code
            if self.processes[filename].returncode == 0:
                self.root.after(0, lambda: self.update_app_status(filename, "stopped"))
                self.root.after(0, lambda: self.status_var.set(f"{filename} completed successfully"))
            else:
                self.root.after(0, lambda: self.update_app_status(filename, "error"))
                self.root.after(0, lambda: self.status_var.set(f"{filename} exited with error"))
                
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to run {filename}: {str(e)}"))
            self.root.after(0, lambda: self.update_app_status(filename, "error"))
            self.root.after(0, lambda: self.status_var.set("Ready"))
            
    def stop_application(self, filename):
        """Stop the specified application"""
        if filename in self.processes:
            try:
                # Terminate the process
                self.processes[filename].terminate()
                self.processes[filename].wait(timeout=5)
                self.update_app_status(filename, "stopped")
                self.status_var.set(f"Stopped {filename}")
            except subprocess.TimeoutExpired:
                # Force kill if terminate didn't work
                self.processes[filename].kill()
                self.processes[filename].wait()
                self.update_app_status(filename, "stopped")
                self.status_var.set(f"Force stopped {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to stop {filename}: {str(e)}")
        else:
            self.status_var.set(f"{filename} is not running")
            
    def update_app_status(self, filename, status):
        """Update the status indicator for an application"""
        # Update color based on status
        if filename in self.status_labels:
            if status == "running":
                self.status_labels[filename].config(fg="#10b981")  # Green
            elif status == "stopped":
                self.status_labels[filename].config(fg="#94a3b8")  # Gray
            elif status == "error":
                self.status_labels[filename].config(fg="#ef4444")  # Red

def main():
    root = tk.Tk()
    app = ApplicationDashboard(root)
    root.mainloop()

if __name__ == "__main__":
    main()