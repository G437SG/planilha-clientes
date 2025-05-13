# Form Exporter

This project is a Python application designed to facilitate the collection and export of form data into both PDF and Excel formats. It provides a user-friendly interface for inputting client and project information, which can then be exported for reporting and documentation purposes.

## Project Structure

The project is organized as follows:

```
form-exporter
├── src
│   ├── main.py               # Entry point of the application, initializes the GUI and handles exports.
│   ├── ui.py                 # Contains user interface components and layout definitions.
│   ├── logic
│   │   ├── app_logic.py      # Manages application state and user input, including export functionality.
│   │   └── export.py         # Handles data formatting and writing to Excel.
│   └── assets
│       └── logo_empresa.png   # Logo image used in exports.
├── requirements.txt           # Lists project dependencies.
├── setup.py                   # Configuration for packaging the application for distribution.
└── README.md                  # Documentation for installation and usage.
```

## Installation

To set up the project, follow these steps:

1. Clone the repository:
   ```
   git clone <repository-url>
   cd form-exporter
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

To run the application, execute the following command:
```
python src/main.py
```

This will launch the GUI where you can input client and project information. Once the data is entered, you can export it to PDF or Excel formats using the provided buttons.

## Exporting Data

The application allows you to export the collected data in two formats:

- **PDF**: Generates a formatted PDF report including all the input data.
- **Excel**: Exports the data to an Excel file, maintaining the same formatting as the PDF.

## Packaging for macOS

To create a .app file for macOS, run the following command:
```
python setup.py py2app
```

This will generate a standalone application that can be distributed and run on macOS systems.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.