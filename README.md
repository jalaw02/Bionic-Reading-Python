# Bionic Reading

Bionic Reading is a Python application that allows you to format text with bionic reading and paragraph shading options and export it to a Word document.

## Features

- Input plain text or import a Word document.
- Apply bionic reading formatting to the text.
- Apply gradient shading to paragraphs.
- Export formatted text to a Word document.

## Requirements

- Python 3.6+
- `tkinter` (usually included with Python)
- `python-docx` library

## Installation

1. **Clone the Repository:**

    ```bash
    git clone https://github.com/yourusername/text-formatter.git
    cd text-formatter
    ```

2. **Create and activate a virtual environment (optional but recommended):**

    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. **Install required libraries:**

    ```bash
    pip install python-docx
    ```

## Usage

1. **Run the Script:**

    ```bash
    python text_formatter.py
    ```

2. **Using the Application:**
    - **Input Text:**
      - Enter text directly into the provided text box, or
      - Click the "Convert to Word Document" button to import a Word document.
    - **Select Formatting Options:**
      - Check the "Bionic Reading" checkbox to apply bionic reading formatting.
      - Check the "Paragraph Shading" checkbox to apply gradient shading to paragraphs.
    - **Convert and Export:**
      - Click the "Convert to Word Document" button.
      - Choose the save location and filename for the new Word document.
      - The application will create and save the formatted document, and show a success message.

## Example

### Input:
```plaintext
This is an example text to demonstrate bionic reading and gradient shading.
```

### Output (Bionic Reading and Shading enabled):
- **Bionic Reading:** Each word is split into bold and normal parts.
- **Gradient Shading:** Paragraphs are shaded with a color gradient.

## Contributing

1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/your-feature`).
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin feature/your-feature`).
5. Create a new Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

### Note

Ensure you have the `tkinter` and `python-docx` libraries installed. The script relies on these libraries for the GUI and Word document manipulation respectively.
