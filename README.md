# Batch Google Hits Inserter

A desktop app for batch Google search using SerpAPI, updating hit counts in Excel files.  
Supports real API use and test/mock mode. Adjustable advanced settings.  
Easy to use for both technical and non-technical users.

---

## Features

- Simple graphical user interface (Tkinter)
- Batch searching using [SerpAPI](https://serpapi.com/)
- Test/mock mode for developing and demonstrating the app without using credits
- Adjustable Google domain, language, and country code ("Avanceret indstillinger")
- Easily set or change your API key from the GUI
- Updates only the "HITS" column in your Excel sheet, preserving all other data
- Saves results as a new file or overwrites the original
- Modular, clean code ready for maintenance and extension

---

## Installation

1. **Clone/download the repository**

2. **Install requirements**
    ```sh
    pip install -r requirements.txt
    ```

3. **Create a `.env` file in the project folder:**
    ```
    SERPAPI_API_KEY=your_api_key_here
    ```

---

## Usage

1. **Start the app**
    ```sh
    python main.py
    ```

2. **Follow the on-screen instructions:**
    - Select your Excel file with keywords.
    - (Optional, for test mode) Select a mock data file.
    - Click "Avanceret indstillinger" to adjust Google domain, language, country, or set your API key.
    - Choose whether to overwrite your file or save as a new one.
    - Click "Start batch" to process.

3. **Review the results:**  
    - Only the "HITS" column is updated in your Excel file.
    - Errors are reported in the GUI and flagged in the result file.

---

## Advanced Options

Click **"Avanceret indstillinger"** in the app to:
- Change Google domain (default: `google.dk`)
- Change language code (`hl`, default: `da`)
- Change country code (`gl`, default: `dk`)
- Edit and save your API key
- Toggle "Testtilstand (MOCK)" on or off

---

## Test/Mock Mode

- Test mode lets you develop, demonstrate, or experiment without using SerpAPI credits.
- In test mode, you must select a mock data file (Excel file with a single column of hit counts).

---

## Making an Executable

To create a standalone Windows `.exe` file:

1. Install [PyInstaller](https://pyinstaller.org/):

    ```sh
    pip install pyinstaller
    ```

2. Build the executable:

    ```sh
    pyinstaller --onefile --windowed gui_app.py
    ```

3. Distribute the `.exe` from the `dist/` folder (along with your `.env` and any example files as needed).

---

## Requirements

See `requirements.txt`, including:
- google-search-results
- python-dotenv
- requests
- pandas
- openpyxl

---


## Author

Steffen Buhl Krøis

