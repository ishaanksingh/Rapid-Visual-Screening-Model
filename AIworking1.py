import tkinter as tk
from tkinter import ttk

def on_yes_no_select(section, subpart, value):
    print(f"Section {section}.{subpart} selected: {value}")

def on_seismic_zone_select(event):
    selected_zone = seismic_zone_var.get()
    print(f"Selected Seismic Zone: {selected_zone}")

def on_occupancy_select(event):
    selected_occupancy = occupancy_var.get()
    print(f"Selected Occupancy Type: {selected_occupancy}")

def on_continue_button_click():
    notebook.select(1)

def validate_radio_selection(radio_group, var):
    if not any(option.get() for option in radio_group):
        error_label = tk.Label(second_tab, text="Please select an option.", fg="red")
        error_label.grid(row=14, column=0, sticky="w")
        root.update_idletasks()
        error_label.after(2000, error_label.destroy)  # Destroy label after 2 seconds

root = tk.Tk()
root.title("Rapid Visual Screening")

# Create StringVars
seismic_zone_var = tk.StringVar()
occupancy_var = tk.StringVar()

# Tab Control
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# Starting Tab
starting_tab = ttk.Frame(notebook)
starting_tab_label = ttk.Label(starting_tab, text="Welcome to the Rapid Visual Screening Model", font=("Arial", 14, "bold"))
starting_tab_label.grid(row=0, column=0, padx=10, pady=10)
continue_button = ttk.Button(starting_tab, text="Continue", command=on_continue_button_click)
continue_button.grid(row=1, column=0, padx=10, pady=10)

# Second Tab
second_tab = ttk.Frame(notebook)
second_tab_heading_label = ttk.Label(second_tab, text="Select the Type of Structure to Perform RVS On", font=("Arial", 14, "bold"))
second_tab_heading_label.grid(row=0, column=0, padx=10, pady=10)

structure_type_label = ttk.Label(second_tab, text="Please select the type of structure for which the RVS is conducted:")
structure_type_label.grid(row=1, column=0, sticky="w", padx=10, pady=5)

structure_type_var = tk.StringVar()
structure_type_concrete = ttk.Radiobutton(
    second_tab, text="Concrete Structure", variable=structure_type_var, value="Concrete Structure"
)
structure_type_rc = ttk.Radiobutton(
    second_tab, text="Rigid Frame Structure (RC)", variable=structure_type_var, value="Rigid Frame Structure (RC)"
)
structure_type_masonry = ttk.Radiobutton(
    second_tab, text="Masonry Structure", variable=structure_type_var, value="Masonry Structure"
)
structure_type_other = ttk.Radiobutton(
    second_tab, text="Other (Please Specify)", variable=structure_type_var, value="Other (Please Specify)"
)

structure_type_concrete.grid(row=2, column=0, sticky="w", padx=10)
structure_type_rc.grid(row=3, column=0, sticky="w", padx=10)
structure_type_masonry.grid(row=4, column=0, sticky="w", padx=10)
structure_type_other.grid(row=5, column=0, sticky="w", padx=10)

other_structure_type_label = ttk.Label(
    second_tab, text="Please specify the structure type if you selected 'Other' above:"
)
other_structure_type_label.grid(row=6, column=0, sticky="w", padx=10, pady=5)
other_structure_type_entry = ttk.Entry(second_tab, width=50)
other_structure_type_entry.grid(row=7, column=0, sticky="w", padx=10, pady=5)

# Add the starting tab to the notebook
notebook.add(starting_tab, text="Welcome")

# Add the second tab to the notebook
notebook.add(second_tab, text="Structure Type")

# Section 1: Siting Issues
def section1_siting_issues():
    section1_label = ttk.Label(second_tab, text="1. Siting Issues")
    section1_label.grid(row=2, column=0, sticky="w", columnspan=3, padx=10, pady=5)

    subpart1_1_label = ttk.Label(second_tab, text="1.1 Building is resting on ground that has failed due to Landslide/Fissures and Liquefaction.")
    subpart1_1_label.grid(row=3, column=0, sticky="w", padx=10, pady=5)

    subpart1_1_var = tk.StringVar()
    subpart1_1_yes = ttk.Radiobutton(
        second_tab,
        text="Yes",
        variable=subpart1_1_var,
        value="Yes",
        command=lambda: validate_radio_selection(
            [subpart1_1_yes, subpart1_1_no], subpart1_1_var
        ),
    )
    subpart1_1_no = ttk.Radiobutton(
        second_tab,
        text="No",
        variable=subpart1_1_var,
        value="No",
        command=lambda: validate_radio_selection(
            [subpart1_1_yes, subpart1_1_no], subpart1_1_var
        ),
    )
    subpart1_1_yes.grid(row=3, column=1)
    subpart1_1_no.grid(row=3, column=2)

section1_siting_issues()

# Run the GUI
root.mainloop()
