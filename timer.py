import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import time
import csv
import json
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path
import os
import subprocess
import win32gui

class Tooltip:
    def __init__(self, widget, text=''):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind('<Enter>', self.show_tooltip)
        self.widget.bind('<Leave>', self.hide_tooltip)

    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20

        # Create top level window
        self.tooltip = tk.Toplevel(self.widget)
        # Remove the window decorations
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = tk.Label(self.tooltip, text=self.text,
                      justify='left',
                      background="#ffffff",
                      relief='solid', borderwidth=1,
                      font=("TkDefaultFont", "8", "normal"))
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class Split:
    def __init__(self, name):
        self.name = name
        self.split_time = None
        self.segment_time = None
        self.best_segment = None
        self.focus_time = 0  # Total focused time
        self.focus_window = None  # Window to track
        self.is_focusing = False  # Currently tracking focus?

class SpeedrunTimerGUI:

    LAST_TEMPLATE_FILE = "last_template_path.txt"

    def __init__(self, root):
        self.root = root
        self.always_on_top = tk.BooleanVar(value=False)  # Track always-on-top state
        self.root.title("Speedrun Timer")
        self.root.configure(bg="white")

        # Timer variables
        self.start_time = None
        self.is_running = False
        self.current_split_index = 0
        self.last_split_time = 0
        self.splits = []
        self.elapsed_time = 0
        self.run_type = "DEFAULT"

        self.create_menu()
        self.create_gui()
        self.update_timer()

        # Try to load last exported template
        last_template = self.get_last_template_path()
        if last_template:
            self.import_run_template(last_template)
        else:
            # Load default template as fallback
            self.load_run_template("RIGID_SCHEDULE")
            self.run_type = "RIGID_SCHEDULE"

    def save_last_template_path(self, file_path):
        """Save the path of the last exported template"""
        try:
            with open(self.LAST_TEMPLATE_FILE, 'w') as f:
                f.write(file_path)
        except Exception as e:
            print(f"Error saving last template path: {str(e)}")

    def get_last_template_path(self):
        """Get the path of the last exported template"""
        try:
            if os.path.exists(self.LAST_TEMPLATE_FILE):
                with open(self.LAST_TEMPLATE_FILE, 'r') as f:
                    path = f.read().strip()
                    if os.path.exists(path):
                        return path
        except Exception as e:
            print(f"Error reading last template path: {str(e)}")
        return None

    def create_menu(self):
        """Create the menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)

        file_menu.add_command(label="Import Run Template", command=self.import_run_template)
        file_menu.add_command(label="Export Run Template", command=self.export_run_template)
        file_menu.add_separator()
        file_menu.add_command(label="Export Times to CSV", command=self.export_times_to_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)


        # Edit Menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Edit Splits", command=self.edit_splits)
        edit_menu.add_separator()
        edit_menu.add_command(label="Save Current Run", command=self.save_current_run)
        edit_menu.add_command(label="Load Current Run", command=self.load_current_run)

        # Preferences Menu
        preferences_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Preferences", menu=preferences_menu)
        preferences_menu.add_checkbutton(
            label="Always on Top",
            variable=self.always_on_top,
            command=self.toggle_always_on_top
        )

    def edit_splits(self):
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Splits")
        edit_window.geometry("600x400")

        # Create Treeview for editing
        edit_tree = ttk.Treeview(edit_window, columns=("Split Name", "Split Time", "Segment Time", "Best Segment"), show="headings")

        for col in ("Split Name", "Split Time", "Segment Time", "Best Segment"):
            edit_tree.heading(col, text=col, anchor="center")
            edit_tree.column(col, width=150, anchor="center")

        # Create button frame
        button_frame = ttk.Frame(edit_window)
        button_frame.pack(fill=tk.X, padx=10, pady=5, side=tk.BOTTOM)

        # 1. Define all helper functions first
        def validate_time_format(time_str):
            if time_str == "":
                return True
            try:
                parts = time_str.split(':')
                if len(parts) != 3:
                    return False
                hours, minutes, seconds = parts
                hours = int(hours)
                minutes = int(minutes)
                seconds = int(seconds)
                return True
            except:
                return False

        def parse_time_to_seconds(time_str):
            if not time_str or time_str == "":
                return None
            try:
                parts = time_str.split(':')
                hours = int(parts[0])
                minutes = int(parts[1])
                seconds = float(parts[2])
                return hours * 3600 + minutes * 60 + seconds
            except:
                return None

        def on_double_click(event):
            region = edit_tree.identify("region", event.x, event.y)
            if region != "cell":
                return

            item = edit_tree.identify_row(event.y)
            if not item:
                return

            column = edit_tree.identify_column(event.x)
            if not column:
                return

            edit_tree.selection_set(item)

            column_index = int(column[1]) - 1
            x, y, w, h = edit_tree.bbox(item, column)

            entry = ttk.Entry(edit_tree, width=20)
            entry.place(x=x, y=y, width=w, height=h)

            current_values = edit_tree.item(item)['values']
            entry.insert(0, current_values[column_index] if current_values[column_index] else "")
            entry.select_range(0, tk.END)
            entry.focus()

            def on_entry_complete(event=None):
                new_value = entry.get()

                if column_index in [1, 2, 3]:
                    if new_value and not validate_time_format(new_value):
                        messagebox.showerror("Error", "Invalid time format. Use HH:MM:SS")
                        return

                current_values = list(edit_tree.item(item)['values'])
                current_values[column_index] = new_value

                edit_tree.item(item, values=current_values)
                entry.destroy()

            entry.bind('<Return>', on_entry_complete)
            entry.bind('<FocusOut>', on_entry_complete)

        def add_split():
            split_name = f"New Split {len(self.splits) + 1}"
            edit_tree.insert("", "end", values=(split_name, "", "", ""))

        def move_up():
            selected = edit_tree.selection()
            if not selected:
                return

            for item in selected:
                idx = edit_tree.index(item)
                if idx > 0:
                    values = edit_tree.item(item)['values']
                    edit_tree.delete(item)
                    edit_tree.insert("", idx-1, values=values)
                    edit_tree.selection_set(edit_tree.get_children()[idx-1])

        def move_down():
            selected = edit_tree.selection()
            if not selected:
                return

            for item in reversed(selected):
                idx = edit_tree.index(item)
                if idx < len(edit_tree.get_children()) - 1:
                    values = edit_tree.item(item)['values']
                    edit_tree.delete(item)
                    edit_tree.insert("", idx+1, values=values)
                    edit_tree.selection_set(edit_tree.get_children()[idx+1])

        def delete_selected():
            selected = edit_tree.selection()
            for item in selected:
                edit_tree.delete(item)

        def save_changes():
            self.splits.clear()
            for item in edit_tree.get_children():
                values = edit_tree.item(item)['values']
                split = Split(values[0])

                if values[1]:  # Split Time
                    try:
                        time_parts = values[1].split(':')
                        hours = int(time_parts[0])
                        minutes = int(time_parts[1])
                        seconds = float(time_parts[2])
                        split.split_time = hours * 3600 + minutes * 60 + seconds
                    except:
                        split.split_time = None

                if values[2]:  # Segment Time
                    try:
                        time_parts = values[2].split(':')
                        hours = int(time_parts[0])
                        minutes = int(time_parts[1])
                        seconds = float(time_parts[2])
                        split.segment_time = hours * 3600 + minutes * 60 + seconds
                    except:
                        split.segment_time = None

                if values[3]:  # Best Segment
                    try:
                        time_parts = values[3].split(':')
                        hours = int(time_parts[0])
                        minutes = int(time_parts[1])
                        seconds = float(time_parts[2])
                        split.best_segment = hours * 3600 + minutes * 60 + seconds
                    except:
                        split.best_segment = None

                self.splits.append(split)

            self.update_splits_display()
            edit_window.destroy()

        # 2. Bind events
        edit_tree.bind('<Double-1>', on_double_click)

        # 3. Populate tree
        for item in edit_tree.get_children():
            edit_tree.delete(item)

        for split in self.splits:
            edit_tree.insert("", "end", values=(
                split.name,
                self.format_time(split.split_time) if split.split_time is not None else "",
                self.format_time(split.segment_time) if split.segment_time is not None else "",
                self.format_time(split.best_segment) if split.best_segment is not None else ""
            ))

        edit_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 4. Create all buttons
        add_button = ttk.Button(button_frame, text="➕", width=3, command=add_split)
        add_button.pack(side=tk.LEFT, padx=2)

        up_button = ttk.Button(button_frame, text="↑", width=3, command=move_up)
        up_button.pack(side=tk.LEFT, padx=2)

        down_button = ttk.Button(button_frame, text="↓", width=3, command=move_down)
        down_button.pack(side=tk.LEFT, padx=2)

        delete_button = ttk.Button(button_frame, text="🗑️", width=3, command=delete_selected)
        delete_button.pack(side=tk.LEFT, padx=2)

        save_button = ttk.Button(button_frame, text="Save Changes", command=save_changes)
        save_button.pack(side=tk.RIGHT, pady=10, padx=10)

    def update_best_segments(self, template_file_path):
        """Update best segments in the template file if they're different"""
        try:
            with open(template_file_path, 'r') as f:
                template_data = json.load(f)

            template_splits = template_data["Current_Template"]["splits"]
            updated = False

            for i, split in enumerate(self.splits):
                if split.segment_time is not None:
                    current_best = template_splits[i].get("best_segment")
                    # Update if the segment time is different from current best
                    if current_best != split.segment_time:
                        template_splits[i]["best_segment"] = split.segment_time
                        updated = True

            if updated:
                with open(template_file_path, 'w') as f:
                    json.dump(template_data, f, indent=4)
                return True
            return False

        except Exception as e:
            print(f"Error updating best segments: {str(e)}")
            return False

    def import_run_template(self, file_path=None):
        """Import a run template from a JSON file"""
        if file_path is None:
            file_path = filedialog.askopenfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                title="Import Run Template"
            )

        if file_path:
            try:
                with open(file_path, 'r') as f:
                    templates = json.load(f)

                self.splits.clear()

                # Set run_type based on the file name
                self.run_type = Path(file_path).stem  # Gets filename without extension

                if isinstance(templates, dict):
                    template_data = templates.get("Current_Template", {})
                    if isinstance(template_data, dict):
                        splits_data = template_data.get("splits", [])
                        for split_data in splits_data:
                            split = Split(split_data["name"])
                            split.best_segment = split_data.get("best_segment")
                            self.splits.append(split)
                    else:
                        # Handle old format for backward compatibility
                        self.splits = [Split(name) for name in template_data]

                self.update_splits_display()
                self.reset_timer()
                if file_path == self.get_last_template_path():
                    print("Last template loaded successfully")
                else:
                    messagebox.showinfo("Success", "Template loaded successfully")

            except Exception as e:
                messagebox.showerror("Error", f"Error loading template: {str(e)}")

    def export_run_template(self):
        """Export current run template to a JSON file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Export Run Template"
        )

        if file_path:
            try:
                template_data = {
                    "Current_Template": {
                        "splits": [
                            {
                                "name": split.name,
                                "best_segment": split.best_segment
                            } for split in self.splits
                        ]
                    }
                }

                with open(file_path, 'w') as f:
                    json.dump(template_data, f, indent=4)

                # Save the path of the exported template
                self.save_last_template_path(file_path)  # Changed this line to use self

                messagebox.showinfo("Success", "Template exported successfully")

            except Exception as e:
                messagebox.showerror("Error", f"Error exporting template: {str(e)}")

    def create_gui(self):
        style = ttk.Style()
        # Use Times New Roman with center alignment for both headings and rows.
        style.configure("Treeview", font=("Segoe UI Variable", 10), background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=("Audrey", 10, "bold"), anchor="center", background="white")

        self.timer_display = tk.Label(
            self.root,
            text="00:00:00",
            font=("Prestage", 36),
            fg="black",
            bg="white"
        )
        self.timer_display.pack(pady=10)

        button_frame = tk.Frame(self.root, bg="white")
        button_frame.pack(pady=5)

        self.start_button = ttk.Button(button_frame, text="Start", command=self.toggle_timer)
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.split_button = ttk.Button(button_frame, text="Split", command=self.hit_split, state=tk.DISABLED)
        self.split_button.pack(side=tk.LEFT, padx=5)
        self.reset_button = ttk.Button(button_frame, text="Reset", command=self.reset_timer)
        self.reset_button.pack(side=tk.LEFT, padx=5)

        splits_frame = tk.Frame(self.root, bg="white")
        splits_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.splits_tree = ttk.Treeview(
            splits_frame,
            columns=("Split Name", "Split Time", "Segment Time", "Best Segment", "Focus Time"),
            show="headings"
        )

        # Set all headings and columns to center
        for col in ("Split Name", "Split Time", "Segment Time", "Best Segment", "Focus Time"):
            self.splits_tree.heading(col, text=col, anchor="center")
            self.splits_tree.column(col, anchor="center", width=120)

        self.splits_tree.pack(fill=tk.BOTH, expand=True)
        self.splits_tree.bind('<Button-1>', self.handle_focus_click)

    def toggle_timer(self):
        if not self.is_running:
            self.start_timer()
        else:
            self.stop_timer()

    def start_timer(self):
        """Start the timer and handle automatic first split if it's wake-up time"""
        self.is_running = True
        current_time = datetime.now().time()

        # Check if this is the first split and if it's wake-up related
        if (self.current_split_index == 0 and
            any(name in self.splits[0].name.lower() for name in ["wake", "get up", "wakeup"])):

            try:
                # Use the hardcoded path to get wake time
                csv_path = r"C:\Users\Kegs\Desktop\fitbit\Data\speedrun_stats.csv"
                df = pd.read_csv(csv_path)

                # Get today's date and wake time
                today = datetime.now().strftime('%Y-%m-%d')
                today_data = df[df['Date'] == today]

                if not today_data.empty:
                    wake_time_str = today_data['Wake Time'].iloc[0]
                    wake_time = datetime.strptime(wake_time_str, '%I:%M %p').time()

                    # Calculate time difference
                    wake_datetime = datetime.combine(datetime.now().date(), wake_time)
                    current_datetime = datetime.combine(datetime.now().date(), current_time)
                    time_diff = (current_datetime - wake_datetime).total_seconds()

                    # Automatically complete first split
                    self.splits[0].split_time = time_diff
                    self.splits[0].segment_time = time_diff
                    self.last_split_time = time_diff
                    self.current_split_index = 1
                    self.update_splits_display()

                    # Important: Set elapsed_time to match the time difference
                    self.elapsed_time = time_diff
                    
                    self.update_splits_display()
                    print(f"Auto-completed first split: {time_diff} seconds since wake-up")
                else:
                    print(f"No wake time data found for today ({today})")

            except Exception as e:
                print(f"Error processing wake time: {str(e)}")

        self.start_time = time.time() - self.elapsed_time
        self.start_button.config(text="Stop")
        self.split_button.config(state=tk.NORMAL)

    def stop_timer(self):
        self.is_running = False
        self.start_button.config(text="Start")
        self.split_button.config(state=tk.DISABLED)

    def reset_timer(self):
        self.stop_timer()
        self.elapsed_time = 0
        self.current_split_index = 0
        self.last_split_time = 0
        self.timer_display.config(text="00:00:00.000")
        self.clear_splits_display()
        self.update_splits_display()

    def update_timer(self):
        if self.is_running:
            self.elapsed_time = time.time() - self.start_time
            self.timer_display.config(text=self.format_time(self.elapsed_time))
            
            # Always update the current split's split_time and segment_time
            if self.current_split_index < len(self.splits):
                current_split = self.splits[self.current_split_index]
                current_split.split_time = self.elapsed_time

                # Added lines to keep segment_time updated in real-time
                # This ensures the "Segment Time" column displays the current segment duration
                if self.current_split_index == 0:
                    current_split.segment_time = self.elapsed_time
                else:
                    current_split.segment_time = (
                        self.elapsed_time - self.splits[self.current_split_index - 1].split_time
                    )

                self.update_splits_display()

        # Keep calling update_timer periodically
        self.root.after(1000, self.update_timer)

    def format_time(self, seconds):
        if seconds is None:
            return ""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)  # Changed to int() to remove decimals
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"  # Removed .3f format

    def format_focus_cell(self, split):
        """
        Return a composite string in the format "FocusTime/FocusPercentage".
        If no focus time is recorded, return "?".
        """
        # If no focus time is recorded, return "?"
        if not split.focus_time:
            return "-"
        
        # Otherwise, format the focus time.
        focus_time_str = self.format_time(split.focus_time)
        
        # Calculate the focus percentage if segment_time is available and > 0.
        if split.segment_time and split.segment_time > 0:
            pct = (split.focus_time / split.segment_time) * 100
            focus_pct_str = f"{pct:.0f}%"
        else:
            focus_pct_str = "0%"
        
        return f"{focus_time_str}/{focus_pct_str}"

    def hit_split(self):
        if self.current_split_index >= len(self.splits):
            return

        split = self.splits[self.current_split_index]
        current_time = self.elapsed_time

        split.split_time = current_time
        split.segment_time = current_time - self.last_split_time

        if split.best_segment is None or split.segment_time < split.best_segment:
            split.best_segment = split.segment_time

        self.last_split_time = current_time
        self.current_split_index += 1

        self.update_splits_display()

        if self.current_split_index >= len(self.splits):
            self.stop_timer()

    def load_run_template(self, template_name):
        try:
            with open('run_templates.json', 'r') as f:
                templates = json.load(f)
                if template_name in templates:
                    self.splits = [Split(name) for name in templates[template_name]]
                    self.update_splits_display()
                    return True
        except FileNotFoundError:
            self.splits = [
                Split("Wake Up"),
                Split("Brush Teeth"),
                Split("Breakfast"),
                Split("Work Start")
            ]
            self.save_run_template(template_name)
            self.update_splits_display()
        return False

    def save_run_template(self, template_name):
        templates = {}
        try:
            with open('run_templates.json', 'r') as f:
                templates = json.load(f)
        except FileNotFoundError:
            pass

        templates[template_name] = [split.name for split in self.splits]

        with open('run_templates.json', 'w') as f:
            json.dump(templates, f, indent=4)

    def clear_splits_display(self):
        for item in self.splits_tree.get_children():
            self.splits_tree.delete(item)

    def update_splits_display(self):
        self.clear_splits_display()
        for i, split in enumerate(self.splits):
            # Optionally, set a background color based on focus percentage if desired.
            bg_color = "white"
            if split.focus_time and split.segment_time:
                percentage = (split.focus_time / split.segment_time) * 100
                bg_color = self.get_focus_color(percentage)
            
            # Get the composite Focus Time string
            focus_cell_text = self.format_focus_cell(split)
            
            item = self.splits_tree.insert("", "end", values=(
                split.name,
                self.format_time(split.split_time) if split.split_time is not None else "",
                self.format_time(split.segment_time) if split.segment_time is not None else "",
                self.format_time(split.best_segment) if split.best_segment is not None else "",
                focus_cell_text
            ))
            
            # Set background color for the row (if needed)
            if split.focus_time and split.segment_time:
                self.splits_tree.tag_configure(f'focus_color_{i}', background=bg_color)
                self.splits_tree.item(item, tags=(f'focus_color_{i}',))

    def export_times_to_csv(self):
        """
        Export current run times to CSV and update best segments if improved.
        Includes 'Focus %' column in the output CSV.
        """
        current_date = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        default_filename = f"speedrun_{self.run_type}_{current_date}.csv"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Export Times to CSV",
            initialfile=default_filename
        )

        if file_path:
            try:
                with open(file_path, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Run Type', self.run_type])
                    writer.writerow(['Date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                    writer.writerow([])
                    # Added "Focus %" column here
                    writer.writerow(['Split Name', 'Split Time', 'Segment Time', 'Best Segment', 'Focus Time', 'Focus %'])

                    for split in self.splits:
                        # Calculate focus percentage if both focus_time and segment_time exist and segment_time > 0
                        focus_pct = ""
                        if split.focus_time and split.segment_time and split.segment_time > 0:
                            pct_value = (split.focus_time / split.segment_time) * 100
                            focus_pct = f"{pct_value:.1f}%"

                        writer.writerow([
                            split.name,
                            self.format_time(split.split_time) if split.split_time is not None else "",
                            self.format_time(split.segment_time) if split.segment_time is not None else "",
                            self.format_time(split.best_segment) if split.best_segment is not None else "",
                            self.format_time(split.focus_time) if split.focus_time else "",
                            focus_pct  # Added Focus %
                        ])

                # Then, ask user if they want to update the template with new best segments
                if messagebox.askyesno("Update Best Segments", "Would you like to update the template with any new best segments?"):
                    template_file = filedialog.askopenfilename(
                        defaultextension=".json",
                        filetypes=[("JSON files", "*.json")],
                        title="Select Template to Update"
                    )
                    if template_file:
                        if self.update_best_segments(template_file):
                            messagebox.showinfo("Success", "Times exported and best segments updated successfully")
                        else:
                            messagebox.showinfo("Success", "Times exported (no new best segments)")
                    else:
                        messagebox.showinfo("Success", "Times exported successfully")
                else:
                    messagebox.showinfo("Success", "Times exported successfully")

            except Exception as e:
                messagebox.showerror("Error", f"Error exporting times: {str(e)}")

    def get_todays_wake_time(self):
        """Read today's wake time from speedrun_stats.csv"""
        try:
            # Hardcoded path to CSV
            csv_path = r"C:\Users\Kegs\Desktop\fitbit\Data\speedrun_stats.csv"

            df = pd.read_csv(csv_path)

            # Get today's date in the same format as CSV
            today = datetime.now().strftime('%Y-%m-%d')
            today_data = df[df['Date'] == today]

            if not today_data.empty:
                wake_time_str = today_data['Wake Time'].iloc[0]  # Format: '06:17 AM'
                wake_time = datetime.strptime(wake_time_str, '%I:%M %p').time()
                return wake_time

            print(f"No data found for today ({today})")
            return None

        except Exception as e:
            print(f"Error reading wake time: {str(e)}")
            return None

    def save_current_run(self):
        """Save the current run state to a JSON file"""
        try:
            current_state = {
                "elapsed_time": self.elapsed_time,
                "current_split_index": self.current_split_index,
                "last_split_time": self.last_split_time,
                "run_type": self.run_type,
                "splits": []
            }

            # Save each split's current state
            for split in self.splits:
                split_data = {
                    "name": split.name,
                    "split_time": split.split_time,
                    "segment_time": split.segment_time,
                    "best_segment": split.best_segment
                }
                current_state["splits"].append(split_data)

            # Save to a dedicated file
            save_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "current_run_state.json")
            with open(save_path, 'w') as f:
                json.dump(current_state, f, indent=4)

            messagebox.showinfo("Success", "Current run saved successfully")

        except Exception as e:
            messagebox.showerror("Error", f"Error saving current run: {str(e)}")

    def load_current_run(self):
        """Load the previously saved run state"""
        try:
            save_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "current_run_state.json")

            if not os.path.exists(save_path):
                messagebox.showwarning("Warning", "No saved run state found")
                return

            with open(save_path, 'r') as f:
                saved_state = json.load(f)

            # Restore timer state
            self.elapsed_time = saved_state["elapsed_time"]
            self.current_split_index = saved_state["current_split_index"]
            self.last_split_time = saved_state["last_split_time"]
            self.run_type = saved_state["run_type"]

            # Restore splits
            self.splits = []
            for split_data in saved_state["splits"]:
                split = Split(split_data["name"])
                split.split_time = split_data["split_time"]
                split.segment_time = split_data["segment_time"]
                split.best_segment = split_data["best_segment"]
                self.splits.append(split)

            # Update display
            self.timer_display.config(text=self.format_time(self.elapsed_time))
            self.update_splits_display()

            messagebox.showinfo("Success", "Run state loaded successfully")

        except Exception as e:
            messagebox.showerror("Error", f"Error loading saved run: {str(e)}")

    def toggle_always_on_top(self):
        """Toggle always-on-top state"""
        self.root.attributes('-topmost', self.always_on_top.get())

    def handle_focus_click(self, event):
        """Handle clicks on the focus button"""
        region = self.splits_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.splits_tree.identify_column(event.x)
            if str(column) == "#5":  # Focus Time column
                item = self.splits_tree.identify_row(event.y)
                if item:
                    index = self.splits_tree.index(item)
                    self.setup_focus_tracking(index)

    def setup_focus_tracking(self, split_index):
        """Setup window tracking for a split"""
        if not self.is_running:
            messagebox.showinfo("Info", "Timer must be running to track focus")
            return

        if split_index != self.current_split_index:
            messagebox.showinfo("Info", "Can only track focus for current split")
            return

        split = self.splits[split_index]

        # If already tracking, stop tracking
        if split.is_focusing:
            split.is_focusing = False
            split.focus_window = None
            messagebox.showinfo("Focus Tracking", "Focus tracking stopped")
            return

        # Ask user to click on the window they want to track
        messagebox.showinfo("Setup Focus Tracking",
            "After clicking OK, click on the window you want to track (you have 3 seconds)")

        self.root.after(3000, lambda: self.capture_window(split_index))

    def capture_window(self, split_index):
        """Capture the currently active window for tracking"""
        try:
            window = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(window)

            split = self.splits[split_index]
            split.focus_window = window_title
            split.is_focusing = True

            messagebox.showinfo("Focus Tracking",
                f"Now tracking window: {window_title}\nFocus time will only count when this window is active")

            # Start the focus checking
            self.check_window_focus()

        except Exception as e:
            messagebox.showerror("Error", f"Error setting up focus tracking: {str(e)}")

    def check_window_focus(self):
        """Check if the tracked window is in focus and update times"""
        if self.is_running and self.current_split_index < len(self.splits):
            current_split = self.splits[self.current_split_index]

            if current_split.is_focusing:
                current_window = win32gui.GetWindowText(win32gui.GetForegroundWindow())

                if current_window == current_split.focus_window:
                    # Update focus time
                    current_split.focus_time += 1
                    self.update_splits_display()

        # Check again in 1 second
        self.root.after(1000, self.check_window_focus)

    def get_focus_color(self, focus_percentage):
        """Return the appropriate color based on focus percentage"""
        color_map = [
            (0, "#FCC0C7"),
            (5.5, "#eebcc7"),
            (11, "#e4bcc4"),
            (16.5, "#dabcc1"),
            (22, "#d0bcbe"),
            (27.5, "#c6bcbb"),
            (33, "#bcbcb8"),
            (38.5, "#b2bcb5"),
            (44, "#a8bcb2"),
            (49.5, "#9ebcaf"),
            (55, "#94bcac"),
            (59.5, "#8abca9"),
            (66, "#80bca6"),
            (71.5, "#76bca3"),
            (77, "#6cbca0"),
            (82.5, "#62bc9d"),
            (88, "#58bc9a"),
            (93.5, "#4EBC97"),
            (99, "#00c62b")
        ]

        # Find the closest percentage match
        closest = min(color_map, key=lambda x: abs(x[0] - focus_percentage))
        return closest[1]


def run_csv_script():
    """Open the CSV script in a new command prompt window"""
    try:
        # Replace with your CSV script path
        csv_script_path = r"C:\Users\Kegs\Desktop\fitbit\fibit.py"

        # Open in new command prompt window
        os.system(f'start cmd /k python "{csv_script_path}"')
        return True

    except Exception as e:
        print(f"Error opening CSV script: {str(e)}")
        return False

def main():
    # First, ask if user wants to update CSV
    initial_root = tk.Tk()
    initial_root.withdraw()  # Hide the empty tkinter window

    should_update = messagebox.askyesno(
        "Update CSV Data",
        "Would you like to update CSV data before starting?\n(This will open the Fitbit data script)"
    )

    initial_root.destroy()  # Destroy the initial window

    if should_update:
        run_csv_script()
        temp_root = tk.Tk()
        temp_root.withdraw()
        if messagebox.askyesno("Continue", "Click Yes when the CSV update is complete to start the timer"):
            temp_root.destroy()
            # Create new root window for main application
            main_root = tk.Tk()
            app = SpeedrunTimerGUI(main_root)
            main_root.mainloop()
        else:
            temp_root.destroy()
    else:
        # Create new root window for main application
        main_root = tk.Tk()
        app = SpeedrunTimerGUI(main_root)
        main_root.mainloop()

if __name__ == "__main__":
    main()