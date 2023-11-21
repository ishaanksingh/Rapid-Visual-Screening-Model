import PySimpleGUI as sg
import subprocess

def create_title_page():
    sg.theme('DarkGrey5')

    layout = [
        [sg.Text('Rapid Visual Screening Model', font=('Helvetica', 20), justification='center')],
        [sg.Text('Created by Ishank Singh, Hemant Singh, Gautam Tayal', font=('Helvetica', 12), justification='center')],
        [sg.Text('Roll Numbers: 2K20/CE/72, 2K20/CE/60, 2K20/CE/66', font=('Helvetica', 12), justification='center')],
        [sg.Multiline('Objectives:\n\n'
                      '1) Identifying the parameters influencing the failure of RC building\n\n'
                      '2) Developing a ML/AI based model to determine the failure of RC Building\n\n'
                      '3) Developing a SQL database of the case studies and extract the most significant parameters using machine learning (ML) algorithms in Python.\n\n'
                      '4) Construct a model to predict the probability of RCC structural failure and train it on the database.\n\n'
                      '5) Augment the database and refine the model\'s output to enhance accuracy.',
                      font=('Helvetica', 14), size=(80, 30), justification='center')],
        [sg.Button('Start', size=(10, 1), font=('Helvetica', 12))]
    ]

    window = sg.Window('Rapid Visual Screening Model', layout, finalize=True)

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == 'Start':
            window.close()
            create_building_selection_page()

def create_building_selection_page():
    sg.theme('DarkGrey5')

    layout = [
        [sg.Text('Select the Type of Building:', font=('Helvetica', 16), justification='center', pad=(0, 20))],
        [sg.Button('RC', size=(15, 2), font=('Helvetica', 12))],
        [sg.Button('Masonry (Coming Soon)', size=(20, 2), font=('Helvetica', 12), disabled=True)]
    ]

    window = sg.Window('Building Selection', layout, finalize=True)

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'RC':
            # Start another Python application (WINDOW 2.PY)
            script_path = "C:\\Users\\ISHANK\\Desktop\\PROJRCT GUI\\WINDOW 2.PY"
            subprocess.run(["python", script_path])
            window.close()  # Close the window when "RC" is clicked

    window.close()

if __name__ == "__main__":
    create_title_page()
