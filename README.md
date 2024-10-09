# Nmap XML to Excel Converter

This Python script parses Nmap XML result files from a specified folder and generates an Excel report for files containing **open ports**. Each XML file with open ports will have its own sheet in the Excel workbook. The script helps in analyzing and reporting open ports across multiple hosts easily.

## Features
- Processes all XML files in a specified folder.
- Filters and includes only ports with an "open" state.
- Generates a single Excel file with individual sheets for each XML file containing open ports.
- Adds a default sheet if no open ports are found in any of the XML files.
- Displays a summary of files containing open ports.

## Prerequisites

Before running the script, make sure you have the following installed:

- Python 3.x
- Required Python libraries:
  - `pandas`
  - `openpyxl`

You can install the necessary dependencies using the following command:

```bash
pip install pandas openpyxl
```

## How to Use

1. **Clone the repository**:

    ```bash
    git clone https://github.com/chathuminda99/nmap-xml-to-excel.git
    cd nmap-xml-to-excel
    ```

2. **Place your Nmap XML files** in a folder.

3. **Run the script**:

    ```bash
    python script.py <folder_path>
    ```

    Replace `<folder_path>` with the path to the folder containing your Nmap XML files.

### Example:

```bash
python script.py /home/user/nmap_results
```

### Output:
- The script will generate an Excel file named `<folder_name>_nmap_results.xlsx`, where `<folder_name>` is the name of the folder you provided. 
- Each sheet in the Excel file corresponds to an XML file that contains open ports. If no open ports are found in any file, a default sheet with a message will be created.

### Sample Output:

```
Processed: scan1.xml (Open ports found)
Processed: scan2.xml (No open ports)
Excel file 'nmap_results_nmap_results.xlsx' created successfully with sheets for files containing open ports.

Files containing open ports:
- scan1.xml
```

## Error Handling

- If the specified folder does not exist, an error message will be displayed: `Error: The folder '<folder_path>' does not exist.`

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

Let me know if you'd like any modifications!
