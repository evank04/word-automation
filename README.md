# Excel to Word Automation Project

This project automates the process of transferring data from multiple Excel spreadsheets into a formatted Word document. The purpose is to keep a digital record of hardcopies.
<p align="center">
  <img src="https://evank04projectbucket.s3.ap-southeast-1.amazonaws.com/Screenshot+2024-12-21+175108.png" alt="Example Output">
</p>

## Features

- **Character Recgonition with AWS Textract**: Extracts handwritten text, and outputs the result into a csv format.
- **Excel Data Extraction**: Reads data from multiple Excel sheets using the `openpyxl` library.
- **Word Document Generation**: Populates a Word document template using the `docxtpl` library.
- **GUI Integration**: User-friendly interface built with `tkinter` for selecting files and initiating the conversion process.
- **Error Reduction**: Ensures consistent formatting and reduces manual data entry errors.
- **Time Efficiency**: Significantly decreases the time required for creating reports by automating the process.

## Dependencies

- Libraries:
  - `openpyxl`
  - `docxtpl`
  - `tkinter`
  - `threading`


## File Structure

- **`parse_data.py`**: Script for processing CSV files and formatting data for Excel.
- **`xl2doc.py`**: Main script for the GUI and data conversion.
- **`Base Template/sat_template_2.docx`**: Word document template with placeholders.
- **`Excel Data`**: Directory containing sample Excel files.
- **`Converted Report`**: Directory for storing generated Word reports.

## Project Outcome

- **Time Savings**: Reduced the time required to transcribe report from ~60 min to ~10 min
- **Error Reduction**: Automation minimized manual entry errors, ensuring higher accuracy.
- **User-Friendly**: Simplified the process for users with no technical background via the GUI.


## Author

**Khuan Jing Jie Evan**
- Diploma in Computer Engineering, Singapore Polytechnic.
- Internship at mVizn.

