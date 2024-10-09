import os
import pandas as pd
import sys
import xml.etree.ElementTree as ET

def nmap_xml_to_excel(folder_path):
    # Check if the folder exists
    if not os.path.isdir(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        return

    # Extract the folder name
    folder_name = os.path.basename(os.path.normpath(folder_path))

    # Create a Pandas Excel writer using openpyxl engine with the folder name included
    excel_file = f'{folder_name}_nmap_results.xlsx'
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')

    # List to store files with open ports
    files_with_open_ports = []
    sheets_created = False  # Flag to check if any sheets were added

    # Iterate over all files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xml'):
            file_path = os.path.join(folder_path, file_name)

            # Parse the XML file
            tree = ET.parse(file_path)
            root = tree.getroot()

            # Extract data from the XML structure
            data = []
            for host in root.findall('host'):
                ip = host.find('address').get('addr') if host.find('address') is not None else 'N/A'
                for port in host.findall('.//port'):
                    state = port.find('state').get('state') if port.find('state') is not None else 'unknown'
                    if state == "open":
                        port_id = port.get('portid')
                        protocol = port.get('protocol')
                        service = port.find('service').get('name') if port.find('service') is not None else 'unknown'
                        data.append([ip, port_id, protocol, state, service])

            # Only proceed if there's open port data
            if data:
                files_with_open_ports.append(file_name)
                df = pd.DataFrame(data, columns=['IP Address', 'Port', 'Protocol', 'State', 'Service'])

                # Use the file name (without extension) as the sheet name, truncated to 31 characters
                sheet_name = os.path.splitext(file_name)[0][:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                sheets_created = True
                print(f"Processed: {file_name} (Open ports found)")
            else:
                print(f"Processed: {file_name} (No open ports)")

    # If no sheets were created, add a default sheet
    if not sheets_created:
        df_empty = pd.DataFrame([["No open ports found in any files."]], columns=["Message"])
        df_empty.to_excel(writer, sheet_name="No_Open_Ports", index=False)
        print("No open ports found in any files. A default sheet has been added.")

    # Save the Excel file
    writer.close()
    print(f"Excel file '{excel_file}' created successfully with sheets for files containing open ports.")

    # Display files that have open ports
    if files_with_open_ports:
        print("\nFiles containing open ports:")
        for file in files_with_open_ports:
            print(f"- {file}")
    else:
        print("\nNo files with open ports found.")

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python script.py <folder_path>")
    else:
        folder_path = sys.argv[1]
        nmap_xml_to_excel(folder_path)
