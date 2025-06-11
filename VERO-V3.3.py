import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from datetime import datetime, timedelta
from PIL import Image, ImageTk

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
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
    file_path = filedialog.askopenfilename(title="Select Cart Cancelled File",
    filetypes=[("CSV Files", "*.csv")]
    )
    if file_path:
        cart_file_label.config(text=os.path.basename(file_path))
        global cart_cancelled_path
        cart_cancelled_path = file_path

def select_sales_file():
    file_path = filedialog.askopenfilename(title="Select Sales Item File",
    filetypes=[("CSV Files", "*.csv")]
    )
    if file_path:
        sales_file_label.config(text=os.path.basename(file_path))
        global sales_item_path
        sales_item_path = file_path
        
def select_output_folder():
    folder_path = filedialog.askdirectory(title="Select Output Folder")
    if folder_path:
        output_folder_label.config(text=os.path.basename(folder_path))
        global output_folder_path
        output_folder_path = folder_path
        
# Main Program Logic - errors first
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
 # Time logic
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
                        df1.at[i, 'Product'] = f"{row1['Product']}====="        # mark rows
                    break
                if row1['Product'] == row2['Product'] and time_diff <= 180 and row1['Kiosk'] != row2['Kiosk']:
                    if mark_option == "delete":
                        if confirm_behavior == "remove":
                            rows_to_remove.append(i)
                        elif confirm_behavior == "mark":
                            df1.at[i, 'Product'] = f"{row1['Product']}====="    # mark rows
                    elif mark_option == "mark":
                        df1.at[i, 'Product'] = f"{row1['Product']}*****"
                    break

    except Exception as e:          # It broke
        progress.stop()
        messagebox.showerror("Error", f"Processing failed: {e}")
        progress["value"] = 0
        return

    df1 = df1.drop(rows_to_remove)

    # Organize columns and specifically destory Kiosk
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
root.geometry("715x550")
root.resizable(False, False)

import tkinter.font as tkfont
default_font = tkfont.Font(family="MS Sans Serif", size=12)
root.option_add("*Font", default_font)

style = ttk.Style(root)      
style.theme_use('classic')        
style.configure("Custom.Horizontal.TProgressbar",
                troughcolor='#c0c0c0',
                background='#000080',
                bordercolor='#808080',
                lightcolor='#A0A0A0',
                darkcolor='#000080',
                thickness=20)


main_frame = tk.Frame(root)
main_frame.pack(fill='both', expand=True)


side_frame = tk.Frame(main_frame, width=200, height=425, bg='#c0c0c0')
side_frame.pack(side='left', anchor='n')
side_frame.pack_propagate(False)

banner_dir = resource_path("Banners")
banner_images = []
banner_index = 0

img = Image.open(resource_path("Banner 1.png"))
max_width = 200
max_height = 425
new_width, new_height = max_width, max_height
img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

tk_img = ImageTk.PhotoImage(img)

side_label = tk.Label(side_frame, image=tk_img, bg='#c0c0c0')
side_label.image = tk_img
side_label.pack(expand=True) 
content_frame = tk.Frame(main_frame, bg='#c0c0c0', padx=10, pady=10)
content_frame.pack(side='left', fill='both', expand=True)



max_width, max_height = 175, 375
box_ratio = max_width / max_height

# find banners - 6 second timer
banner_files = sorted(
    f for f in os.listdir(banner_dir) if f.startswith("Banner") and f.endswith(".png")
)

for banner_file in banner_files:
    banner_path = os.path.join(banner_dir, banner_file)
    try:
        img = Image.open(banner_path)
        img_ratio = img.width / img.height
        if img_ratio > box_ratio:
            new_width = max_width
            new_height = int(max_width / img_ratio)
        else:
            new_height = max_height
            new_width = int(max_height * img_ratio)
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        tk_img = ImageTk.PhotoImage(img)
        banner_images.append(tk_img)
    except Exception as e:
        print(f"Failed loading {banner_path}: {e}")

if banner_images:
    side_label.configure(image=banner_images[0])
    side_label.image = banner_images[0]

def rotate_banner():
    global banner_index
    if banner_images:
        banner_index = (banner_index + 1) % len(banner_images)
        side_label.configure(image=banner_images[banner_index])
        side_label.image = banner_images[banner_index]
    root.after(6000, rotate_banner)

root.after(6000, rotate_banner)



file_select_frame = tk.Frame(content_frame, bg='#c0c0c0')
file_select_frame.pack(fill='x', pady=5)

# Cart Cancelled
cart_frame = tk.Frame(file_select_frame, bg='#c0c0c0')
cart_frame.pack(side='left', padx=5, fill='y')

cart_file_label = tk.Label(cart_frame, text="Cart Cancelled File: Not Selected", bg='#c0c0c0', anchor='w')
cart_file_label.pack()

cart_file_button = tk.Button(cart_frame, text="Select Cart Cancelled File", command=select_cart_file)
cart_file_button.pack()
cart_file_button.config(cursor="hand2")

# Sales Item
sales_frame = tk.Frame(file_select_frame, bg='#c0c0c0')
sales_frame.pack(side='left', padx=5, fill='y')

sales_file_label = tk.Label(sales_frame, text="Sales Item File: Not Selected", bg='#c0c0c0', anchor='w')
sales_file_label.pack()

sales_file_button = tk.Button(sales_frame, text="Select Sales Item File", command=select_sales_file)
sales_file_button.pack()
sales_file_button.config(cursor="hand2")

# Output Folder
output_folder_label = tk.Label(content_frame, text="Output Folder: Not Selected", bg='#c0c0c0', anchor='w')
output_folder_label.pack(fill='y', pady=2)

output_folder_button = tk.Button(content_frame, text="Select Output Folder", command=select_output_folder)
output_folder_button.pack(pady=2)
output_folder_button.config(relief="raised", bd=2, highlightthickness=0)
output_folder_button.config(cursor="hand2")

main_frame.config(highlightbackground="black", highlightthickness=1)
content_frame.config(highlightbackground="black", highlightthickness=1)
side_frame.config(highlightbackground="black", highlightthickness=1)
# ====time adjustments=====
time_adjustment_var = tk.StringVar(value="no_adjustment")

time_adjust_label = tk.Label(content_frame, text="Time Adjustment Options:", bg='#c0c0c0', anchor='w')
time_adjust_label.pack(fill='y', pady=(10, 2))

frame_time_adjust = tk.Frame(content_frame, bg='#c0c0c0')
frame_time_adjust.pack(fill='y', pady=2)

no_adjustment_radio = tk.Radiobutton(frame_time_adjust, text="No Adjustment", variable=time_adjustment_var, value="no_adjustment", bg='#c0c0c0')
no_adjustment_radio.pack(side='left', padx=5)

add_time_radio = tk.Radiobutton(frame_time_adjust, text="Add Time", variable=time_adjustment_var, value="add", bg='#c0c0c0')
add_time_radio.pack(side='left', padx=5)

subtract_time_radio = tk.Radiobutton(frame_time_adjust, text="Subtract Time", variable=time_adjustment_var, value="subtract", bg='#c0c0c0')
subtract_time_radio.pack(side='left', padx=5)

time_adjustment_label = tk.Label(content_frame, text="Enter Time Adjustment (HH:MM:SS or HHMMSS):", bg='#c0c0c0', anchor='w')
time_adjustment_label.pack(fill='y', pady=2)

time_adjustment_entry = tk.Entry(content_frame)
time_adjustment_entry.pack(fill='y', pady=2)


confirm_removal_var = tk.StringVar(value="remove")  # Default is to remove


no_adjustment_radio.config(indicatoron=1)
add_time_radio.config(indicatoron=1)
subtract_time_radio.config(indicatoron=1)


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
# progress bar
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

# Bunch of configs for click inputs
padding_width = 20
padding_height = 100

cart_file_button.config(relief="raised", bd=2)
sales_file_button.config(relief="raised", bd=2)
output_folder_button.config(relief="raised", bd=2)


cart_frame.config(relief="sunken", bd=2)
sales_frame.config(relief="sunken", bd=2)
file_select_frame.config(relief="sunken", bd=2)

for btn in [cart_file_button, sales_file_button, output_folder_button, run_button]:
    btn.config(relief="raised", bd=2, cursor="hand2")


for frame in [cart_frame, sales_frame, file_select_frame, content_frame, side_frame, main_frame]:
    frame.config(relief="sunken", bd=2, highlightbackground="black", highlightthickness=1)


for rb in [no_adjustment_radio, add_time_radio, subtract_time_radio, delete_radio, mark_radio, delete_after_radio, keep_original_radio, remove_matches_radio, mark_matches_radio]:
    rb.config(indicatoron=1, bg='#c0c0c0')


root.geometry(f"{width + padding_width}x{height + padding_height}")


root.mainloop()
