import pandas as pd
import PySimpleGUI as sg
import os
import subprocess
import time

# Create a custom progress bar layout
def progress_bar_layout(bar_size):
    return [
        [sg.Button(' ', size=(1, 1), key=f'-CUBE-{i}-', button_color=('black', 'black'), pad=(0, 0)) for i in range(bar_size)]
    ]

def create_database():
    area = values['area']

    # Connect to the database and create a database for the specified area
    # (Add your code here for connecting to the database)

    message_elem.update(value="Database for area '" + area + "' created successfully!")
    window['-OPEN_BUTTON-'].update(disabled=False)

def open_database(file_path):
    # Open the Excel file with the default program
    os.startfile(file_path)

    message_elem.update(value="Database for area opened successfully!")
    window['-NEXT_BUTTON-'].update(disabled=False)

def show_selected_parameters(data):
    # Implement the action for "Show Selected Parameters"
    message_elem.update(value="Show Selected Parameters clicked!")

    # Start the 'timepass.py' script
    script_path = os.path.join(os.path.dirname(__file__), 'timepass.py')
    subprocess.run(["python", script_path])

def final_step():
    # Implement the action for "Go to Final Step"
    message_elem.update(value="Go to Final Step clicked!")

    # Start the 'WINDOW3.py' script
    script_path = os.path.join(os.path.dirname(__file__), 'WINDOW 3.PY')
    subprocess.run(["python", script_path])

# Main GUI layout
layout = [
    [sg.Text("Enter the area in which the building is located:")],
    [sg.InputText(key='area')],
    [sg.Button("Submit", key='-SUBMIT_BUTTON-')],
    [sg.Frame('', progress_bar_layout(10), key='-PROGRESS_BAR_FRAME-', visible=False)],
    [sg.Text("Click 'Submit' to create a database for the above area.", key='-MESSAGE_LABEL-')],
    [sg.Button("Open Complete Database", key='-OPEN_BUTTON-', disabled=True, bind_return_key=True),
     sg.Button("Show Selected Parameters", key='-NEXT_BUTTON-', disabled=True),
     sg.Button("Go to Final Step", key='-FINAL_STEP_BUTTON-', disabled=True)]
]

window = sg.Window("Create Database", layout, finalize=True)

message_elem = window['-MESSAGE_LABEL-']
progress_bar_frame = window['-PROGRESS_BAR_FRAME-']
cube_elements = [window[f'-CUBE-{i}-'] for i in range(10)]  # Adjust the number of cubes as needed
values = None

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        break

    if event == '-SUBMIT_BUTTON-':
        window['-OPEN_BUTTON-'].update(disabled=True)
        window['-NEXT_BUTTON-'].update(disabled=True)
        window['-FINAL_STEP_BUTTON-'].update(disabled=True)

        # Show the progress bar
        progress_bar_frame.update(visible=True)
        window.refresh()

        for i in range(len(cube_elements)):
            time.sleep(0.5)  # Adjust the sleep duration for your desired speed
            cube_elements[i].update(button_color=('green', 'green'))
            window.refresh()

        create_database()

        window['-OPEN_BUTTON-'].update(disabled=False)

    elif event == '-OPEN_BUTTON-':
        # You need to replace this line with your actual file path.
        # For example, you can use sg.popup_get_file to let the user select the file.
        file_path = "C:\\Users\\ISHANK\\Desktop\\PROJRCT GUI\\DATABASEAREA.xlsx"
        open_database(file_path)

    elif event == '-NEXT_BUTTON-':
        show_selected_parameters(None)
        window['-FINAL_STEP_BUTTON-'].update(disabled=False)

    elif event == '-FINAL_STEP_BUTTON-':
        final_step()

window.close()
