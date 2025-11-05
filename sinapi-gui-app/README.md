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

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.