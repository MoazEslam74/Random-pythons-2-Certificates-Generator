<p align="center">
  <img src="RP_Certificates_Generator.ico" alt="Centered image" width="300">
</p>

# 🎓 Smart Certificate Generator

A robust, Python-based GUI application designed to automate the mass generation of certificates. By leveraging MS Word templates (`.docx`) and COM automation, this tool accurately maps dynamic data to templates and exports them as high-quality PDFs. 

## ✨ Features

*   **🌍 Bilingual Interface:** Seamlessly toggle between English and Arabic UI.
*   **🏷️ Dynamic Tagging:** Add custom tags (e.g., `name`, `date`, `manager`) on the fly to match your Word template placeholders.
*   **🔒 Data Locking:** Lock constant fields (like Instructor Name or Date) to avoid redundant typing across multiple certificates.
*   **👁️ In-App Preview:** Instantly preview a sample certificate as an image within the app before generating the entire batch.
*   **📑 PDF Merging:** Option to automatically merge all generated certificates into a single, print-ready PDF file.
*   **📊 Progress Tracking:** Real-time progress bar and status indicators to monitor bulk generation tasks.
*   **🧵 Non-blocking UI:** Heavy Word COM operations are handled via background threads to ensure the interface remains responsive.

## 🛠️ Prerequisites

*   **OS:** Windows (Required for `win32com` Word automation).
*   **Software:** Microsoft Word must be installed on the machine.
*   **Python Version:** 3.8+

## 📦 Installation

1. Clone this repository or download the source code.
2. Install the required Python dependencies:
   ```bash
   pip install docxtpl pywin32 PyPDF2 PyMuPDF Pillow

## 🚀 Usage

1. Prepare Template: Create a Word document (`.docx`) and use Jinja2 syntax for placeholders (e.g., `{{name}}`, `{{date}}`).

2. Launch the App.

3. Configure Paths: Select your Word template and the desired output directory.

4. Set Tags: Add the exact tag names used in your template via the "+ Add Tag" button.

5. Enter Data: Type the data for each certificate. Use the 🔒 icon to lock repeating data. Press `Enter` to add it to the queue.

6. Preview & Generate: Click **Preview Sample** to verify alignment, then click **Generate Certificates** to start the batch process.

<video src="InfoClip.mp4" autoplay loop muted playsinline width="100%"></video>

## 🏗️ Built With
* [Tkinter](https://docs.python.org/3/library/tkinter.html) - GUI Framework

* [docxtpl](https://docxtpl.readthedocs.io/) - Word Template rendering

* [pywin32](https://github.com/mhammond/pywin32) - COM API for Word-to-PDF conversion

* [PyPDF2](https://pypdf2.readthedocs.io/) - PDF Merging

* [PyMuPDF](https://pymupdf.readthedocs.io/) (fitz) & [Pillow](https://python-pillow.org/) - In-app visual preview
