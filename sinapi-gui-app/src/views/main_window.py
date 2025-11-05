from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QVBoxLayout, QLabel, QComboBox, QDateEdit, QRadioButton, QCheckBox, QPushButton, QHBoxLayout, QGroupBox
from src.controllers.sinapi_controller import SinapiController

class MainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SINAPI GUI App")
        self.setGeometry(100, 100, 400, 300)
        
        self.layout = QVBoxLayout()
        
        # Dropdown for selecting SINAPI, SICRO, ORSE
        self.dropdown_label = QLabel("Select Type:")
        self.dropdown = QComboBox()
        self.dropdown.addItems(["SINAPI", "SICRO", "ORSE"])
        
        # Date input
        self.date_label = QLabel("Select Date:")
        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDate(QtCore.QDate.currentDate())
        
        # Radio buttons for SINAPI types
        self.radio_group = QGroupBox("Select SINAPI Type")
        self.radio_layout = QVBoxLayout()
        self.radio_ambos = QRadioButton("Ambos")
        self.radio_desonerado = QRadioButton("Desonerado")
        self.radio_nao_desonerado = QRadioButton("Não Desonerado")
        self.radio_layout.addWidget(self.radio_ambos)
        self.radio_layout.addWidget(self.radio_desonerado)
        self.radio_layout.addWidget(self.radio_nao_desonerado)
        self.radio_group.setLayout(self.radio_layout)
        
        # Checkboxes for Brazilian states
        self.state_checkboxes = {}
        self.states = ["AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"]
        for state in self.states:
            checkbox = QCheckBox(state)
            self.state_checkboxes[state] = checkbox
            self.layout.addWidget(checkbox)
        
        # Buttons
        self.concluir_button = QPushButton("Concluir")
        self.plus_one_button = QPushButton("+1")
        
        self.layout.addWidget(self.dropdown_label)
        self.layout.addWidget(self.dropdown)
        self.layout.addWidget(self.date_label)
        self.layout.addWidget(self.date_input)
        self.layout.addWidget(self.radio_group)
        self.layout.addWidget(self.concluir_button)
        self.layout.addWidget(self.plus_one_button)
        
        self.setLayout(self.layout)
        
        # Connect buttons to their functions
        self.concluir_button.clicked.connect(self.on_concluir_clicked)
        self.plus_one_button.clicked.connect(self.on_plus_one_clicked)

    def on_concluir_clicked(self):
        selected_type = self.dropdown.currentText()
        selected_date = self.date_input.date()
        selected_year = selected_date.year()
        selected_month = selected_date.month()
        
        selected_states = [state for state, checkbox in self.state_checkboxes.items() if checkbox.isChecked()]
        
        controller = SinapiController()
        controller.execute_sinapi(selected_year, selected_month, selected_type, selected_states)

    def on_plus_one_clicked(self):
        # Logic for +1 button can be implemented here
        pass