# Replace Hidden Font In PPTX
PPTX files openend through Google Slides often have a new font embedded ("Noto Sans Symbols"), and you have will receive an error every time you try to save. This lightweight Python program will try to open the pptx as a ZIP file, locate all xml files, and replaces all the found font strings in these files. Finally, a copy of the pptx is saved in the same filder.

&ensp;<img width="502" height="282" alt="image" src="https://github.com/user-attachments/assets/43a496d1-5a44-487f-ac20-dc419f5d57fc" />

## Features
* **GUI Interface:** No command line required; includes a file browser and status pop-ups.
* **Persistent Memory:** Automatically remembers your last used "Search" and "Replace" strings.
* **Deep Replacement:** Scans all XML components of the PPTX, including slides, notes, and masters.
* **Safe Execution:** Creates a new file with a `_modified` suffix rather than modifying the original.

## How It Works
Modern `.pptx` files are actually compressed archives containing XML. This tool:
1.  Creates a temporary `.zip` copy of your presentation.
2.  Iterates through every internal `.xml` file.
3.  Replaces every instance of your target string.
4.  Re-packages the files and saves them back as a `.pptx` with a new filename.

## Installation
1.  **Requirement:** Ensure you have [Python 3.x](https://www.python.org/downloads/) installed.
2.  **No Dependencies:** This script only uses built-in libraries (`tkinter`, `zipfile`, `json`).
3.  **Download:** Save the script (e.g., `pptx_replacer.py`) to your computer.

## Usage
0. Find the string to replace ("Noto Sans Symbols"):
   
&ensp;<img width="410" height="355" alt="image" src="https://github.com/user-attachments/assets/f211d8c4-ee82-4ffc-abdd-108057e172fe" />


2.  Run the script:
    ```bash
    python pptx_replacer.py
    ```
3.  **Select File:** Click **Browse** to choose your `.pptx` file.
4.  **Search/Replace:** Enter the strings you wish to swap (e.g., Search: `Noto Sans Symbols`, Replace: `Verdana`).
5.  **Run:** Click **Run Replacement**. 
6.  The new file will appear in the same folder as the original.

## ⚠️ Notes
* **Formatting:** You can also use the tool to replace plain text in the XML (text on the slide!). If a word in a slide has mixed formatting (e.g., the first half is **bold** and the second isn't), the XML may split the string, preventing a match.
* **Backups:** While the script creates a new file, it is always good practice to keep a backup of your original data.

## License
This project is open-source and free to use.
