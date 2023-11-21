
import requests
import pandas as pd
from bs4 import BeautifulSoup

def get_html_content(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"Error fetching content from {url}: {e}")
        return None

def check_parameters(html_content):
    parameters_to_check = [
        "Building is resting on ground that has failed due to Landslide/Fissures and Liquefaction.",
        "Building is tilted.",
        # ... (include all other parameters)
    ]

    matched_parameters = []

    soup = BeautifulSoup(html_content, 'html.parser')
    page_text = soup.get_text()

    for parameter in parameters_to_check:
        if parameter.lower() in page_text.lower():
            matched_parameters.append(parameter)

    return matched_parameters

def main():
    # Specify the path to your Excel file
    excel_file_path = 'C:/Users/ISHANK/Desktop/pyrthon/DATABASEAREA.xlsx'

    # Specify the CSV file to store matched parameters
    output_csv_path = 'C:/Users/ISHANK/Desktop/pyrthon/output.csv'

    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(excel_file_path)

        # Print information about the DataFrame
        print("\nDataFrame Information:")
        print(df.info())

        # Print the first few rows of the DataFrame
        print("\nFirst Few Rows of DataFrame:")
        print(df.head())

        # Assume the column containing URLs is named 'URL'
        urls = df['URL']

        all_matched_parameters = []

        for url in urls:
            html_content = get_html_content(url)

            if html_content:
                print(f"\nExtracted Information from {url}:")

                # Check parameters
                matched_parameters = check_parameters(html_content)

                if matched_parameters:
                    print("Matched Parameters:")
                    for parameter in matched_parameters:
                        print(parameter)
                    all_matched_parameters.append({
                        'URL': url,
                        'Matched Parameters': ', '.join(matched_parameters)
                    })

        # Save matched parameters to CSV file
        if all_matched_parameters:
            df_result = pd.DataFrame(all_matched_parameters)
            print("\nDataFrame with Matched Parameters:")
            print(df_result)

            # Try saving to CSV with a different path
            df_result.to_csv(output_csv_path, index=False)
            print(f"\nMatched parameters saved to '{output_csv_path}'")

    except FileNotFoundError:
        print(f"Error: File not found at {excel_file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
