import pandas as pd
import PySimpleGUI as sg

# Read the Excel file into a Pandas DataFrame
df = pd.read_excel('C:\\Users\\ISHANK\\Desktop\\PROJRCT GUI\\dbsolve.xlsx')

# Filter the DataFrame to only include rows where the "Attendance%" column is greater than 35
filtered_df = df[df["%"] > 35]

# Create a list of lists for the table data
table_data = [filtered_df.columns.tolist()] + filtered_df.values.tolist()

# Define the PySimpleGUI layout with a table
layout = [
    [sg.Table(values=table_data,
              headings=filtered_df.columns.tolist(),
              auto_size_columns=False,
              justification='right',
              display_row_numbers=False,
              num_rows=min(25, len(filtered_df)),
              enable_events=True,
              key="-TABLE-",
              col_widths=25)],
    [sg.Button("OK"), sg.Button("Exit")]
]

# Create the window with an increased size
window = sg.Window("Prominent Parameters", layout, size=(1200, 600))

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == "Exit":
        break

    if event == "OK":
        selected_row_index = values["-TABLE-"][0]
        selected_row = filtered_df.iloc[selected_row_index]
        sg.popup("Selected Student: ", selected_row["Tag"])

# Close the window
window.close()
