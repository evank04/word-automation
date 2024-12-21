# Excel to Word Automation Project

This project automates the process of transferring data from multiple Excel spreadsheets into a formatted Word document. It provides a graphical user interface (GUI) for ease of use and ensures consistent and accurate data formatting. The project is ideal for scenarios requiring repetitive data entry and document generation, such as generating Site Acceptance Test (SAT) reports.

![image alt][https://evank04projectbucket.s3.ap-southeast-1.amazonaws.com/Screenshot+2024-12-21+175108.png]
## Features

- **Excel Data Extraction**: Reads data from multiple Excel sheets using the `openpyxl` library.
- **Word Document Generation**: Populates a Word document template using the `docxtpl` library.
- **GUI Integration**: User-friendly interface built with `tkinter` for selecting files and initiating the conversion process.
- **Error Reduction**: Ensures consistent formatting and reduces manual data entry errors.
- **Time Efficiency**: Significantly decreases the time required for creating reports by automating the process.

## Requirements

- Python 3.8+
- Libraries:
  - `openpyxl`
  - `docxtpl`
  - `tkinter`
  - `threading`
- Dependencies can be installed using:
  ```bash
  pip install openpyxl docxtpl
  ```

## How to Use

1. Clone this repository or download the project files.
2. Ensure the required Python libraries are installed.
3. Run the script:
   ```bash
   python xl-doc_convert.py
   ```
4. Use the GUI to:
   - Enter the crane name.
   - Select the Excel file to be processed.
   - Click the **Convert** button to generate the Word document.
5. The output Word document will be saved in the `Converted Report` directory with a timestamped filename.

## File Structure

- **`parse_data.py`**: Script for processing CSV files and formatting data for Excel.
- **`xl-doc_convert.py`**: Main script for the GUI and data conversion.
- **`Base Template/sat_template_2.docx`**: Word document template with placeholders.
- **`Excel Data`**: Directory containing sample Excel files.
- **`Converted Report`**: Directory for storing generated Word reports.

## Future Enhancements

- Implement advanced error handling for invalid inputs.
- Add support for additional file formats (e.g., PDFs).
- Integrate Optical Character Recognition (OCR) for automated data extraction from handwritten or printed documents.
- Enhance the GUI with more customization options and visual feedback.

## Author

**Evan Khuan Jing Jie**
- Diploma in Computer Engineering, Singapore Polytechnic.
- Internship at mVizn.

## License

This project is open-source and available under the MIT License.
