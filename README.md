# panel_kayit-forKivancTekstil-
# Panel Kayıt – Production Barcode Logging System

## Project Overview

This project is a lightweight GUI-based application designed for real-time barcode logging in manufacturing environments. It is actively used in the production line of **Kıvanç Tekstil**, where it automates the process of tracking electrical panel components. The application uses barcode input with automatic timestamping, stores records in Excel files, and allows live record monitoring via a graphical interface.

The system is built in Python using Tkinter and is tailored for single-computer, offline-safe usage scenarios where rapid deployment and ease of use are prioritized.

---

## Technologies Used

- **Python 3.x** – Main programming language
- **Tkinter** – Graphical user interface
- **openpyxl** – Excel file management
- **SQLite** – (Not used in current version but can be easily integrated)
- **datetime** – Timestamping functionality
- **threading** – Timeout-based input control
- **subprocess / os / platform** – OS-level command execution

---

## Core Functionalities

### 1. Barcode Input and Auto-Logging
The application consists of two input fields:
- **Bara Barcode** (e.g., main rail)
- **Panel Barcode** (e.g., attached component)

Upon entering both, the system logs the entry along with the current timestamp into a daily Excel file.

### 2. GUI Interface with Treeview Log
Logged entries are listed in a tabular view using `ttk.Treeview`. This allows the user to monitor entries in real-time and provides visual confirmation for every action.

### 3. Timeout Reset Mechanism
If the user enters a barcode and leaves the system idle, it resets the input fields after a set timeout (12 seconds by default), allowing for uninterrupted production flow.

### 4. Deletion and Record Management
Users can select one or more rows from the GUI and delete them both from the interface and from the Excel file automatically, ensuring data consistency.

### 5. Excel Export with Auto-Creation
Each day, a new Excel file is created under a predefined directory. The file is automatically named based on the current date and stored in a local folder (e.g., `D:/BarkodKayit/03_08_2025.xlsx`).

### 6. Executable Conversion
This application has also been compiled to an `.exe` executable using PyInstaller for direct deployment on Windows machines in the factory.

---


## Screenshot
Main user interface with barcode input and real-time logging table:

<img width="897" height="377" alt="image" src="https://github.com/user-attachments/assets/2821094f-8bcb-48cc-8364-a2ac93172a2c" />
