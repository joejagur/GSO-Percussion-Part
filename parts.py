import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Open file dialog for user to select the Excel file
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
    
    if not file_path:
        print("No file selected.")
        return
    
    songs = setList(file_path)
    
    # Show a song selection dialog with a listbox
    weekly = select_songs(songs)
    
    if not weekly:
        print("No songs selected.")
        return

    df = pd.read_excel(file_path, sheet_name=weekly)

    parts = set()
    for key in df:
        for row in df[key].iterrows():
            if pd.isna(row[1].iloc[0]) or "Perc" in row[1].iloc[0].title() or "Cymbal" in row[1].iloc[0].title():
                part = row[1].iloc[1].capitalize()
                part = fixVarient(part)
                parts.add(part)
            else:
                part = row[1].iloc[0].capitalize().split(", ")
                if len(part) == 2:
                    parts.add(part[1].capitalize())
                part = fixVarient(part[0])
                parts.add(part)

    print("Total Instruments needed:", len(parts))
    parts = list(parts)
    parts.sort()
    
    # Write the parts to a text file
    write_parts_to_file(parts, weekly, songs, "Parts.txt")

    print_parts(parts)

def fixVarient(input):
    input = input.title()
    if "Sus" in input:
        return "Suspended Cymbal"
    elif "Bass" in input:
        return "Bass Drum"
    elif "Snare" in input:
        return "Snare Drum"
    elif "Crash" in input:
        return "Crash Cymbal"
    elif "Glock" in input:
        return "Glockenspiel"
    elif "Tam" in input or "Gong" in input:
        return "Gong"
    elif "Mark Tree" in input:
        return "Windchimes"
    elif "Chimes" in input:
        return "Tubular Bells"
    elif "Vib" in input and "phone" in input:
        return "Vibraphone"
    elif "Chinese" in input or "Big" in input:
        return "Chinese Bass Drum"
    else:
        return input

def print_parts(parts):
    for part in parts:
        print(part)

def write_parts_to_file(parts, songsSelected, allSongs, filename):
    """Write the sorted list of parts to a text file."""
    with open(filename, "w") as f:
        f.write(f"Total Instruments needed: {len(parts)}\n")
        f.write("In Songs: \n")
        for song in songsSelected: 
            f.write(f"{allSongs[song]}\n")
        f.write("\n")
        for part in parts:
            f.write(f"{part}\n")
    print(f"\nParts list written to {filename}")

def setList(file_path):
    xl = pd.ExcelFile(file_path)
    return xl.sheet_names


def select_songs(songs):
    def toggle_select_all():
        is_selected = not select_all_var.get()
        print(f"Checkbox state (1 when checked, 0 when unchecked): {is_selected}")
        if is_selected:
            listbox.select_set(0, tk.END)
        else:
            listbox.select_clear(0, tk.END)
        select_all_var.set(is_selected)
        listbox.event_generate("<<ListboxSelect>>")

    root = tk.Tk()
    root.title("Parts")

    listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, height=10)
    for i, song in enumerate(songs):
        listbox.insert(tk.END, f"{i}: {song}")
    listbox.pack(padx=10, pady=10)

    select_all_var = tk.IntVar()
    select_all_checkbox = tk.Checkbutton(root, text="Select All", variable=select_all_var, command=toggle_select_all)
    select_all_checkbox.pack(pady=5)

    selected_indices = []
    
    def on_select():
        nonlocal selected_indices
        selected_indices = list(listbox.curselection())  # Ensure selected indices are captured as a list
        
        # If no items are selected, warn the user
        if not selected_indices:
            messagebox.showwarning("No Selection", "No songs were selected.")
        else:
            print(f"Selected indices: {selected_indices}")
            root.quit()  # Close the window
            select_button = tk.Button(root, text="Select", command=on_select)
    
    select_button = tk.Button(root, text="Select", command=on_select)
    select_button.pack(pady=5)
    

    root.mainloop()
    
    #root.destroy()

    if selected_indices:
        return [int(idx) for idx in selected_indices]
    return None
main()
