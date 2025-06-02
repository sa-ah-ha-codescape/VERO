import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os
import sys
from datetime import datetime, timedelta
from PIL import Image, ImageTk


def resource_path(relative_path):
    """ Get the absolute path to resource, works for dev and PyInstaller """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Global
cart_cancelled_path = ""
sales_item_path = ""
output_folder_path = ""
mark_option = "untouched" 


def get_desktop_folder():
    if os.name == 'nt':  #Windows(TRASH)
        user_profile = os.environ.get("USERPROFILE", "")
        desktop_folder = os.path.join(user_profile, 'Desktop', 'Reports Sorted')
    else:  #Linux(best) or MacOS(Second Worst)
        desktop_folder = os.path.expanduser("~/Desktop/Reports Sorted")
    return desktop_folder

def str_to_timedelta(time_str):
    try:
        if len(time_str) == 8:  #HH:MM:SS format
            return datetime.strptime(time_str, "%H:%M:%S") - datetime(1900, 1, 1)
        elif len(time_str) == 6:  #HHMMSS format
            return datetime.strptime(time_str, "%H%M%S") - datetime(1900, 1, 1)
        else:
            raise ValueError("Invalid time format")
    except Exception as e:
        print(f"Error parsing time: {e}")
        return None

def adjust_time(current_time, diff, operation):
    if operation == 'add':
        return current_time + diff
    elif operation == 'subtract':
        return current_time - diff
    else:
        return current_time

#File Select Functions
def select_cart_file():
    global cart_cancelled_path
    cart_cancelled_path = filedialog.askopenfilename(title="Select Cart Cancelled CSV", filetypes=[("CSV files", "*.csv")])
    cart_file_label.config(text=f"Cart Cancelled File: {cart_cancelled_path if cart_cancelled_path else 'Not Selected'}")

def select_sales_file():
    global sales_item_path
    sales_item_path = filedialog.askopenfilename(title="Select Sales Item CSV", filetypes=[("CSV files", "*.csv")])
    sales_file_label.config(text=f"Sales Item File: {sales_item_path if sales_item_path else 'Not Selected'}")

def select_output_folder():
    global output_folder_path
    output_folder_path = filedialog.askdirectory(title="Select Output Folder")
    output_folder_label.config(text=f"Output Folder: {output_folder_path if output_folder_path else 'Not Selected'}")

# Main Program Logic
def run_script_gui():
    if not cart_cancelled_path or not sales_item_path or not output_folder_path:
        messagebox.showerror("Error", "Please select all required files and output folder.")
        return
    
    progress["value"] = 0
    progress.start()
    root.update()

    try:
        df1 = pd.read_csv(cart_cancelled_path)
        df2 = pd.read_csv(sales_item_path)
    except Exception as e:
        progress.stop()
        messagebox.showerror("Error", f"Failed to read files: {e}")
        return

    time_adjustment_option = time_adjustment_var.get()

    if time_adjustment_option != "no_adjustment":
        corrected_time_str = time_adjustment_entry.get()
        corrected_time = str_to_timedelta(corrected_time_str)
        if corrected_time is None:
            progress.stop()
            messagebox.showerror("Error", "Invalid time format. Please enter in HH:MM:SS or HHMMSS format.")
            return
        operation = "add" if time_adjustment_option == "add" else "subtract"

        try:
            df1['TransactionDate'] = pd.to_datetime(df1['TransactionDate'], errors='coerce', utc=True)
            df2['MachineLocalTime'] = pd.to_datetime(df2['MachineLocalTime'], errors='coerce', utc=True)

            df1['TransactionDate'] = df1['TransactionDate'].apply(lambda x: adjust_time(x, corrected_time, operation))
            df2['MachineLocalTime'] = df2['MachineLocalTime'].apply(lambda x: adjust_time(x, corrected_time, operation))
        except Exception as e:
            progress.stop()
            messagebox.showerror("Error", f"Failed to adjust time: {e}")
            return

    mark_option = mark_option_var.get()

    df1['TransactionDate'] = pd.to_datetime(df1['TransactionDate'], errors='coerce')
    df2['MachineLocalTime'] = pd.to_datetime(df2['MachineLocalTime'], errors='coerce')

    progress["maximum"] = len(df1)
    progress["value"] = 0
    root.update_idletasks()

    rows_to_remove = []
    try:
        
        confirm_behavior = confirm_removal_var.get()

        for i, row1 in df1.iterrows():
            progress["value"] = i + 1
            root.update_idletasks()
            for j, row2 in df2.iterrows():
                time_diff = abs((row1['TransactionDate'] - row2['MachineLocalTime']).total_seconds())
                if row1['Product'] == row2['Product'] and row1['Kiosk'] == row2['Kiosk'] and time_diff <= 180:
                    if confirm_behavior == "remove":
                        rows_to_remove.append(i)
                    elif confirm_behavior == "mark":
                        df1.at[i, 'Product'] = f"{row1['Product']}====="
                    break
                if row1['Product'] == row2['Product'] and time_diff <= 180 and row1['Kiosk'] != row2['Kiosk']:
                    if mark_option == "delete":
                        if confirm_behavior == "remove":
                            rows_to_remove.append(i)
                        elif confirm_behavior == "mark":
                            df1.at[i, 'Product'] = f"{row1['Product']}====="
                    elif mark_option == "mark":
                        df1.at[i, 'Product'] = f"{row1['Product']}*****"
                    break

    except Exception as e:
        progress.stop()
        messagebox.showerror("Error", f"Processing failed: {e}")
        progress["value"] = 0
        return

    df1 = df1.drop(rows_to_remove)

    # Organize columns; remove Kiosk
    df2 = df2[['Product', 'Price', 'MachineLocalTime', 'Kiosk']]
    df1 = df1[['Product', 'ItemPrice', 'Quantity', 'TransactionDate', 'Kiosk']]

    df1 = df1.sort_values(by='TransactionDate', ascending=False)
    df2 = df2.sort_values(by='MachineLocalTime', ascending=False)

    df1 = df1.drop(columns=['Kiosk'])
    df2 = df2.drop(columns=['Kiosk'])

    delete_originals_choice = delete_originals_var.get()
    if delete_originals_choice == "delete":
        try:
            os.remove(cart_cancelled_path)
            os.remove(sales_item_path)
            print("Original files deleted.")
        except Exception as e:
            print(f"Failed to delete original files: {e}")
    else:
        print("Original files retained.")

    file1_name = 'Cart Cancelled Sorted_Done.csv'
    file2_name = 'Sales Item by Location Done.csv'

    try:
        df1.to_csv(os.path.join(output_folder_path, file1_name), index=False)
        df2.to_csv(os.path.join(output_folder_path, file2_name), index=False)
    except Exception as e:
        progress.stop()
        messagebox.showerror("Error", f"Failed to save files: {e}")
        return

    progress.stop()
    progress["value"] = 0
    messagebox.showinfo("Success", f"Files saved to: {output_folder_path}")

# ====GUI Setup===

root = tk.Tk()
root.title("Report Wizard")
root.configure(bg='#c0c0c0')
root.geometry("550x525")

style = ttk.Style(root)      
style.theme_use('clam')        
style.configure("Custom.Horizontal.TProgressbar",
                troughcolor='#c0c0c0',
                background='#000080',
                bordercolor='#808080',
                lightcolor='#A0A0A0',
                darkcolor='#000080',
                thickness=20)


main_frame = tk.Frame(root)
main_frame.pack(fill='both', expand=True)


side_frame = tk.Frame(main_frame, width=250, height=550, bg='#c0c0c0')
side_frame.pack(side='left', anchor='w',fill='y')
side_frame.pack_propagate(True)


img = Image.open(resource_path("SideBanner.png"))

max_width, max_height = 175, 375
img_ratio = img.width / img.height  
box_ratio = max_width / max_height  
style.theme_use('clam')

if img_ratio > box_ratio:
    new_width = max_width
    new_height = int(max_width / img_ratio)
else:
    new_height = max_height
    new_width = int(max_height * img_ratio)

img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
tk_img = ImageTk.PhotoImage(img)

side_label = tk.Label(side_frame, image=tk_img, bg='#c0c0c0')
side_label.image = tk_img
side_label.pack(expand=True) 


content_frame = tk.Frame(main_frame, bg='#c0c0c0')
content_frame.pack(side='left', fill='both', expand=False)



content_frame = tk.Frame(main_frame, bg='#c0c0c0', padx=10, pady=10)
content_frame.pack(side='left', fill='both', expand=True)


cart_file_label = tk.Label(content_frame, text="Cart Cancelled File: Not Selected", bg='#c0c0c0', anchor='w')
cart_file_label.pack(fill='x', pady=2)

cart_file_button = tk.Button(content_frame, text="Select Cart Cancelled File", command=select_cart_file)
cart_file_button.pack(pady=2, anchor='w')

sales_file_label = tk.Label(content_frame, text="Sales Item File: Not Selected", bg='#c0c0c0', anchor='w')
sales_file_label.pack(fill='x', pady=2)

sales_file_button = tk.Button(content_frame, text="Select Sales Item File", command=select_sales_file)
sales_file_button.pack(pady=2, anchor='w')

output_folder_label = tk.Label(content_frame, text="Output Folder: Not Selected", bg='#c0c0c0', anchor='w')
output_folder_label.pack(fill='x', pady=2)

output_folder_button = tk.Button(content_frame, text="Select Output Folder", command=select_output_folder)
output_folder_button.pack(pady=2, anchor='w')


time_adjustment_var = tk.StringVar(value="no_adjustment")

time_adjust_label = tk.Label(content_frame, text="Time Adjustment Options:", bg='#c0c0c0', anchor='w')
time_adjust_label.pack(fill='x', pady=(10, 2))

frame_time_adjust = tk.Frame(content_frame, bg='#c0c0c0')
frame_time_adjust.pack(fill='x', pady=2)

no_adjustment_radio = tk.Radiobutton(frame_time_adjust, text="No Adjustment", variable=time_adjustment_var, value="no_adjustment", bg='#c0c0c0')
no_adjustment_radio.pack(side='left', padx=5)

add_time_radio = tk.Radiobutton(frame_time_adjust, text="Add Time", variable=time_adjustment_var, value="add", bg='#c0c0c0')
add_time_radio.pack(side='left', padx=5)

subtract_time_radio = tk.Radiobutton(frame_time_adjust, text="Subtract Time", variable=time_adjustment_var, value="subtract", bg='#c0c0c0')
subtract_time_radio.pack(side='left', padx=5)

time_adjustment_label = tk.Label(content_frame, text="Enter Time Adjustment (HH:MM:SS or HHMMSS):", bg='#c0c0c0', anchor='w')
time_adjustment_label.pack(fill='x', pady=2)

time_adjustment_entry = tk.Entry(content_frame)
time_adjustment_entry.pack(fill='x', pady=2)


confirm_removal_var = tk.StringVar(value="remove")  # Default is to remove




mark_option_var = tk.StringVar(value="mark")

mark_option_label = tk.Label(content_frame, text="Flagged Rows Handling:", bg='#c0c0c0', anchor='w')
mark_option_label.pack(fill='x', pady=(10, 2))

frame_mark_option = tk.Frame(content_frame, bg='#c0c0c0')
frame_mark_option.pack(fill='x', pady=2)

delete_radio = tk.Radiobutton(frame_mark_option, text="Delete False-positives", variable=mark_option_var, value="delete", bg='#c0c0c0')
delete_radio.pack(side='left', padx=5)

mark_radio = tk.Radiobutton(frame_mark_option, text="Mark False-positives", variable=mark_option_var, value="mark", bg='#c0c0c0')
mark_radio.pack(side='left', padx=5)


delete_originals_var = tk.StringVar(value="delete")

delete_originals_label = tk.Label(content_frame, text="Original Files:", bg='#c0c0c0', anchor='w')
delete_originals_label.pack(fill='x', pady=(10, 2))

frame_delete_originals = tk.Frame(content_frame, bg='#c0c0c0')
frame_delete_originals.pack(fill='x', pady=2)

delete_after_radio = tk.Radiobutton(frame_delete_originals, text="Delete After Processing", variable=delete_originals_var, value="delete", bg='#c0c0c0')
delete_after_radio.pack(side='left', padx=5)

keep_original_radio = tk.Radiobutton(frame_delete_originals, text="Keep Original Files", variable=delete_originals_var, value="keep", bg='#c0c0c0')
keep_original_radio.pack(side='left', padx=5)

style.configure("Custom.Horizontal.TProgressbar",
    troughcolor='#c0c0c0',   
    background='#000080',    
    bordercolor='#808080',   
    lightcolor='#A0A0A0',    
    darkcolor='#000080',     
    thickness=20            
)
progress = ttk.Progressbar(content_frame, style="Custom.Horizontal.TProgressbar",
                           orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10, anchor='w')

run_button = tk.Button(content_frame, text="Run Script", command=run_script_gui)
run_button.pack(pady=10, anchor='center')

confirm_removal_label = tk.Label(content_frame, text="Confirm Matches Handling:", bg='#c0c0c0', anchor='w')
confirm_removal_label.pack(fill='x', pady=(10, 2))

frame_confirm_removal = tk.Frame(content_frame, bg='#c0c0c0')
frame_confirm_removal.pack(fill='x', pady=2)

remove_matches_radio = tk.Radiobutton(frame_confirm_removal, text="Remove Matches", variable=confirm_removal_var, value="remove", bg='#c0c0c0')
remove_matches_radio.pack(side='left', padx=5)

mark_matches_radio = tk.Radiobutton(frame_confirm_removal, text="Mark Matches Only", variable=confirm_removal_var, value="mark", bg='#c0c0c0')
mark_matches_radio.pack(side='left', padx=5)


root.update()  
width = root.winfo_width()
height = root.winfo_height()


padding_width = 20
padding_height = 100

root.geometry(f"{width + padding_width}x{height + padding_height}")


root.mainloop()
