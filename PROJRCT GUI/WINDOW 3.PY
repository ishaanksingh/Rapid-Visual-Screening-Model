import pandas as pd
import PySimpleGUI as sg

# Read the Excel file into a Pandas DataFrame
df = pd.read_excel('C:\\Users\\ISHANK\\Desktop\\PROJRCT GUI\\dbsolve.xlsx')

# Filter the DataFrame to only include rows where the "Attendance%" column is greater than 35
filtered_df = df[df["%"] > 35]

# Create a PySimpleGUI layout with checkboxes next to each student name
layout = [
    [sg.Text("Parameters Priority > 35%")],
    *[[sg.Checkbox(row['Tag'], default=False, key=f'checkbox_{index}')] for index, row in filtered_df.iterrows()],
    [sg.Button("OK"), sg.Button("Exit")]
]

# Define the PySimpleGUI window
window = sg.Window("Prominent Parameters", layout)

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == "Exit":
        break

    if event == "OK":
        total_checkboxes = len(filtered_df)
        selected_students = [row['Tag'] for index, row in filtered_df.iterrows() if values[f'checkbox_{index}']]
        selected_count = len(selected_students)

        probability_of_failure = (selected_count / total_checkboxes) * 100 if total_checkboxes > 0 else 0

        sg.popup("Selected Parameters:", ', '.join(selected_students))

        if probability_of_failure < 30:
            message = "Low Probability of Failure"
            color = 'green'
        elif 30 <= probability_of_failure < 60:
            message = "Moderate Probability of Failure"
            color = 'orange'
        else:
            message = "High Probability of Failure"
            color = 'red'

        sg.popup(f"Probability of Failure: {probability_of_failure:.2f}%", message, background_color=color)

# Close the window
window.close()
