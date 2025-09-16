# PDF Converter Tool 📄➡️🔄 v0.0.3
A simple tool created with Python that converts PDFs files to Word, Excel, PowerPoint ant TXT.


## 🚀 Features

- **Conversion:** Supports PDF to TXT conversion and text extraction for Word, Excel, and PowerPoint formats.
- **Graphical Interface:** Simple and intuitive user interface built with `Tkinter`.
- **Open Source:** The code is completely open and available for modification and reuse.


## ⚙️ Installation

1.  Clone the repository:
    ```bash
    git clone [https://github.com/matteocheccacci/pdf-converter-tool.git](https://github.com/matteocheccacci/pdf-converter-tool.git)
    ```
2.  Navigate to the project folder:
    ```bash
    cd pdf-converter-tool
    ```
3.  Install the necessary dependencies:
    ```bash
    pip install -r requirements.txt
    ```


## 🖥️ How to use

1.  Run the main script:
    ```bash
    python pdf_converter.py
    ```
2.  Click "Select PDF" and select your file.
3.  Select the desired output format from the drop-down menu.
4.  Click "Convert" and choose the destination folder for the converted file.


## ⚠️ Known Issues

- **Icon Display:** The icon in the title bar may not be displayed correctly on all operating systems or Python versions due to limitations with `tkinter.iconbitmap()` method.
- **Corrupted Excel/PowerPoint Files:** The tool performs a basic text-only conversion for Excel and PowerPoint. As a result, the converted files (`.xlsx` and `.pptx`) may appear corrupted or damage, as they lack the necessary internal structure to properly display the data.
- **Non_Text PDF Conversion:** The tool's core logic is based on text extraction. This meand PDFs containing only images (e.g., scanned documents) cannot be converted in to editable text. The output will be an empty file.
- **Complex Formatting:** The text-only conversion for Word, Excel and PowerPoint files will result in the loss of all original formatting, tables and images. The output is a plain text file saved with the chosen extension.


## 🤝 Contributions

Contributions are welcome! If you'd like to improve the project, feel free to open a **pull request** or report a **bug** by filing an issue.


## 📄 License

This project is distributed under the **MIT** license. For more details, see the `LICENSE` file.

---