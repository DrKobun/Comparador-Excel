# SINAPI GUI Application

This project is a graphical user interface (GUI) application for generating and opening SINAPI links based on user-selected criteria. The application allows users to select a type (SINAPI, SICRO, ORSE), input a date, choose between different types of SINAPI data, and select Brazilian states to generate the appropriate links.

## Project Structure

```
sinapi-gui-app
├── src
│   ├── main.py                # Entry point of the application
│   ├── ui.py                  # User interface definitions
│   ├── sinapi.py              # Code for generating and opening SINAPI links
│   ├── controllers
│   │   └── sinapi_controller.py # Logic for handling user interactions
│   ├── views
│   │   └── main_window.py      # Layout and behavior of the main window
│   └── resources
│       └── states.py          # List of Brazilian states
├── tests
│   └── test_sinapi.py         # Unit tests for functionality
├── requirements.txt            # External libraries required for the project
├── .gitignore                  # Files and directories to ignore by version control
└── README.md                   # Documentation for the project
```

## Setup Instructions

1. Clone the repository:
   ```
   git clone <repository-url>
   cd sinapi-gui-app
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python src/main.py
   ```

2. Use the dropdown to select "SINAPI", "SICRO", or "ORSE".

3. Input the desired date.

4. Select the type of SINAPI data using the radio buttons ("ambos", "Desonerado", "Não Desonerado").

5. Check the boxes for the Brazilian states you want to include.

6. Click "Concluir" to generate and open the links based on your selections.

7. A clickable link for more information is available in the SINAPI section, which opens a web page.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Excel Formatting Scripts

### organizar_excel.py

This script iterates through all Excel files in a specified directory and saves each sheet as a new Excel file. For any sheet that has "EQP" in its name, it will delete columns C, D, E, F, G, H, and I.

This script requires the following libraries:
- pandas
- openpyxl

### formatar_aninhados.py

This script automates the formatting of specific Excel spreadsheets. It iterates over all Excel files in the 'Arquivos-SINAPI-SICRO-ORSE/aninhar' folder on the user's desktop.

- For sheets starting with 'SINA-SIN', it performs a series of formatting operations, including deleting columns and rows, and auto-fitting column widths.
- For sheets starting with 'SINA-INS', it performs a different set of formatting operations.
- For any sheet that has "EQP" in its name, it will delete columns C, D, E, F, G, H, and I.

This script requires the following libraries:
- xlwings