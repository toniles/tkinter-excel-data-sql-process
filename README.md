# tkinter-excel-data-sql-process 📊

[![Python 3.7+](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: Unlicense](https://img.shields.io/badge/License-Unlicense-green.svg)](https://unlicense.org/)

**🚀 Data processing and database updating tool!**

**Data Processor** is a Python script that streamlines the process of updating database records from external data files. While designed for a specific use case, it serves as an excellent model for building similar tools for your company's data management needs.

## 🎯 Key Features

- **📈 Processes structured data files**
- **🔍 Validates data** before database operations
- **🗃️ Performs database operations** based on input data
- **⚠️ Robust error handling** and reporting
- **📊 Real-time progress tracking** with GUI
- **🧵 Multi-threaded processing** for improved performance

## 🚀 Quick Start Guide

### 1️⃣ Prerequisites

- **Python** 3.7 or higher
- Required libraries: `pandas`, `pyodbc`, `tkinter`, `tkinterdnd2`
- Access to the target database

### 2️⃣ Setup

1. Clone the repository or download the script:
   ```bash
   git clone https://github.com/your-repo/data-processor.git
   cd data-processor
   ```

2. Install required dependencies:
   ```bash
   pip install pandas pyodbc tkinterdnd2
   ```

3. Update the database connection string in the script with your details.

### 3️⃣ Running the Script

Execute the script in Python:

```bash
python data_processor.py
```

## 💡 How It Works

1. The script provides a GUI for selecting an input file.
2. It reads the file and validates the data.
3. The data is then processed in two phases:
   - Test phase: Checks data validity without making changes.
   - Process phase: Performs database operations with the input data.
4. Progress is displayed in real-time via the GUI.
5. Results, including successes and errors, are shown in the GUI and saved to a log file.

## 🛠️ Configuration

- The script is designed to work with a specific database structure and input file format.
- Modify the data validation and processing logic to fit your specific needs.

## ⚠️ Note on Sensitive Information

This public version has had sensitive information such as server addresses, database names, and credentials removed. When adapting this script for your use, ensure to replace these with your actual connection details.

## 📄 License

This is free and unencumbered software released into the public domain. See the [Unlicense](https://unlicense.org/) for more details.

```
This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or distribute this 
software, either in source code form or as a compiled binary, for any purpose, 
commercial or non-commercial, and by any means.

In jurisdictions that recognize copyright laws, the author or authors of this 
software dedicate any and all copyright interest in the software to the public 
domain. We make this dedication for the benefit of the public at large and to 
the detriment of our heirs and successors. We intend this dedication to be an 
overt act of relinquishment in perpetuity of all present and future rights to 
this software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN 
ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
```

## 📢 Feedback & Support

While this is a model script and not actively maintained, feel free to use it as inspiration for your own projects. If you have questions about adapting this for your needs, consider opening a discussion in the repository.
