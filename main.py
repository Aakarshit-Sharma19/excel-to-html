import pandas as pd
import os
from pathlib import Path
import webbrowser

def generate_html_file(folder_name, data, name_column_index, phone_column_index, group_size=1000):
    # Split data into groups based on group size
    groups = [data[i:i + group_size] for i in range(0, len(data), group_size)]

    for group_index, group_data in enumerate(groups):
        # Create an HTML string to store the links
        html_content = "<html><body>"

        # Iterate over each row in the group
        for index, row in group_data.iterrows():
            name = row[name_column_index - 1]
            phone_number = row[phone_column_index - 1]

            # Generate the links for call, message, and WhatsApp
            call_link = f"<a href='tel:{phone_number}'>Call</a>"
            message_link = f"<a href='sms:{phone_number}'>Message</a>"
            whatsapp_link = f"<a href='https://api.whatsapp.com/send?phone={phone_number}'>WhatsApp</a>"

            # Create a table row with the person's index, name, and links
            html_content += f"<p><b>Index:</b> {index}</p>"
            html_content += f"<p><b>Name:</b> {name}</p>"
            html_content += f"<p><b>Call:</b> {call_link}</p>"
            html_content += f"<p><b>Message:</b> {message_link}</p>"
            html_content += f"<p><b>WhatsApp:</b> {whatsapp_link}</p>"
            html_content += "<hr>"

        html_content += "</body></html>"

        # Save the HTML content to a file
        file_name = f"{folder_name}/contacts_group_{group_index + 1}.html"
        with open(file_name, "w") as file:
            file.write(html_content)

        print(f"HTML file '{file_name}' generated successfully.")

def main():
    # Get the file path from the user
    file_path = input("Enter the path to the Excel file: ")

    # Read the Excel file into a pandas DataFrame
    try:
        data = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error: {e}")
        return

    # Display the available column indexes to the user
    print("Available columns:")
    for i, column in enumerate(data.columns, start=1):
        print(f"{i}. {column}")

    # Get the column indexes for names and phone numbers from the user
    name_column_index = int(input("Enter the column index for names: "))
    phone_column_index = int(input("Enter the column index for phone numbers: "))

    folder_name = Path(file_path).name + '_output'
    os.mkdir(folder_name)
    # Generate the HTML files
    generate_html_file(folder_name, data, name_column_index, phone_column_index)

    # Open the first HTML file in the default web browser
    webbrowser.open(f"{folder_name}/contacts_group_1.html")

if __name__ == "__main__":
    main()
