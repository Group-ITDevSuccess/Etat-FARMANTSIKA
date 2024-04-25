import datetime
import json
import sys
import qdarkstyle
from PyQt5.QtWidgets import (
    QWidget,
    QPushButton,
    QApplication,
    QListWidget,
    QLabel,
    QHBoxLayout,
    QLineEdit,
    QComboBox,
    QTableWidget,
    QVBoxLayout,
    QTableWidgetItem, QMessageBox, QTabWidget,
)
from PyQt5.QtCore import QTimer, QDateTime, Qt, QTime


from entity.export import exportDataFrameEncaissement
from entity.query import getDateInSQLServer


class WinForm(QWidget):
    def __init__(self, parent=None):
        super(WinForm, self).__init__(parent)
        self.timer = None
        self.table_historique = None
        self.validate_button = None
        self.add_row_button = None
        self.table_destinataire = None
        self.tab_historique = None
        self.tab_destinataire = None
        self.tab_widget = None
        self.endBtn = None
        self.startBtn = None
        self.label = None
        self.listFile = None
        self.day_combo = None
        self.day_label = None
        self.target_second_edit = None
        self.target_minute_edit = None
        self.target_hour_edit = None
        self.target_time_label = None
        self.setFixedSize(627, 600)
        self.setWindowTitle('ID Motors')
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Recherche dans l'historique")  # Search in history

        self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
        self.setupUi()

        self.load_destination_from_json()
        self.load_historique_from_json()

        self.target_time = None
        # self.target_day = None
        self.data_destination = []
        self.data_historique = []

    def setupUi(self):
        self.target_time_label = QLabel('Heure:')
        self.target_hour_edit = QLineEdit('9')
        self.target_hour_edit.setFixedWidth(50)
        self.target_minute_edit = QLineEdit('0')
        self.target_minute_edit.setFixedWidth(50)
        self.target_second_edit = QLineEdit('0')
        self.target_second_edit.setFixedWidth(50)

        # self.day_label = QLabel('Jour:')
        # self.day_combo = QComboBox()
        #  self.day_combo.addItems(['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'])

        self.listFile = QListWidget()
        self.label = QLabel('Etat automatique de la liste des articles')
        self.label.setAlignment(Qt.AlignCenter)  # Align label text to center
        self.startBtn = QPushButton('Start')
        self.endBtn = QPushButton('Stop')

        # Création d'un QTabWidget pour gérer les onglets
        self.tab_widget = QTabWidget()

        # Création des onglets "Destinataire" et "Historique"
        self.tab_destinataire = QWidget()
        self.tab_historique = QWidget()

        self.tab_widget.addTab(self.tab_destinataire, "Destinataire")
        self.tab_widget.addTab(self.tab_historique, "Historique")

        # Tableau pour l'onglet "Destinataire"
        self.table_destinataire = QTableWidget()
        self.table_destinataire.setColumnCount(5)  # Increased to 5 for the new 'Status' column
        self.table_destinataire.setHorizontalHeaderLabels(['Id', 'Nom', 'Email', 'Status', 'Action'])
        self.table_destinataire.setColumnWidth(0, 20)
        self.table_destinataire.setColumnWidth(1, 150)
        self.table_destinataire.setColumnWidth(2, 250)
        self.table_destinataire.setColumnWidth(3, 100)
        self.table_destinataire.setColumnWidth(4, 50)

        self.add_row_button = QPushButton('Ajouter')
        self.add_row_button.clicked.connect(self.add_row_to_table)

        self.validate_button = QPushButton('Valider')
        self.validate_button.clicked.connect(self.save_destination_to_json)

        # Layout pour l'onglet "Destinataire"
        layout_destinataire = QVBoxLayout(self.tab_destinataire)
        layout_destinataire.addWidget(self.table_destinataire)  # Ajoutez le tableau à cet onglet
        layout_destinataire.addWidget(self.add_row_button)
        layout_destinataire.addWidget(self.validate_button)

        # Création d'un deuxième tableau pour l'onglet "Historique"
        self.table_historique = QTableWidget()
        self.table_historique.setColumnCount(6)  # Nombre de colonnes pour l'historique
        self.table_historique.setHorizontalHeaderLabels(['N°', 'Date', 'Destinataire', 'Email', 'Status', 'Action'])
        self.table_historique.setColumnWidth(0, 5)
        self.table_historique.setColumnWidth(1, 115)
        self.table_historique.setColumnWidth(2, 100)
        self.table_historique.setColumnWidth(3, 225)
        self.table_historique.setColumnWidth(4, 50)
        self.table_historique.setColumnWidth(5, 60)
        layout_historique = QVBoxLayout(self.tab_historique)
        # layout_historique.addWidget(self.search_bar)
        layout_historique.addWidget(self.table_historique)

        layout = QVBoxLayout()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_countdown)

        # Target time input layout
        target_time_layout = QHBoxLayout()
        target_time_layout.addWidget(self.target_time_label)
        target_time_layout.addWidget(self.target_hour_edit)
        target_time_layout.addWidget(QLabel('Minute: '))
        target_time_layout.addWidget(self.target_minute_edit)
        target_time_layout.addWidget(QLabel('Second: '))
        target_time_layout.addWidget(self.target_second_edit)
        # target_time_layout.addWidget(self.day_label)
        # target_time_layout.addWidget(self.day_combo)
        target_time_layout.addWidget(self.startBtn)  # Ajouter le bouton "Start" après les champs Heure
        target_time_layout.addWidget(self.endBtn)  # Ajouter le bouton "Stop" après le bouton "Start"

        layout.addLayout(target_time_layout)
        layout.addWidget(self.label)
        layout.addWidget(self.tab_widget)

        self.startBtn.clicked.connect(self.start_countdown)
        self.endBtn.clicked.connect(self.end_timer)

        self.table_destinataire.setColumnHidden(0, True)
        self.table_historique.setColumnHidden(0, True)
        # self.search_bar.textChanged.connect(self.filter_historique_table)

        self.setLayout(layout)


    def filter_historique_table(self, search_text):
        self.table_historique.setRowCount(0)  # Clear the table before filtering

        for item in self.data_historique:
            if search_text.lower() in item['Destinataire'].lower() or search_text.lower() in item['Email'].lower():
                # Add row to the table for matching items
                row_position = self.table_historique.rowCount()
                self.table_historique.insertRow(row_position)

                self.table_historique.setItem(row_position, 0, QTableWidgetItem(str(item['No'])))
                self.table_historique.setItem(row_position, 1, QTableWidgetItem(str(item['DateTime'])))
                self.table_historique.setItem(row_position, 2, QTableWidgetItem(str(item['Destinataire'])))
                self.table_historique.setItem(row_position, 3, QTableWidgetItem(str(item['Email'])))
                self.table_historique.setItem(row_position, 4,
                                              QTableWidgetItem("Success" if item['Status'] else "Echec"))
                if not item['Status']:
                    resend_button = QPushButton('^')
                    resend_button.clicked.connect(lambda _, row=row_position: self.resend_row_from_table(row))
                    self.table_historique.setCellWidget(row_position, 5, resend_button)

    def filter_historique_table_realtime(self):
        current_search_text = self.search_bar.text()
        self.filter_historique_table(current_search_text)

    def add_row_to_table(self):
        row_position = self.table_destinataire.rowCount()
        self.table_destinataire.insertRow(row_position)

        id_item = QTableWidgetItem(str(row_position + 1))  # Définition de la clé 'Id'
        self.table_destinataire.setItem(row_position, 0, id_item)

        name_item = QTableWidgetItem('')
        self.table_destinataire.setItem(row_position, 1, name_item)

        email_item = QTableWidgetItem('')
        self.table_destinataire.setItem(row_position, 2, email_item)

        # Create a combo box for the Status column
        status_combo = QComboBox()
        status_combo.addItems(['Active', 'Inactive'])
        status_combo.setCurrentIndex(0)  # Set default to Active

        self.table_destinataire.setCellWidget(row_position, 3, status_combo)

        action_item = QTableWidgetItem('')
        self.table_destinataire.setItem(row_position, 4, action_item)

        delete_button = QPushButton('X')
        delete_button.clicked.connect(lambda _, row=row_position: self.delete_row_from_table(row))
        self.table_destinataire.setCellWidget(row_position, 4, delete_button)

        # Update data_destination dictionary with the selected status
        self.data_destination.append({'Id': str(row_position + 1), 'Status': 'Active'})  # Default to Active

    def delete_row_from_table(self, row):
        confirm_delete = QMessageBox.question(self, 'Confirmation', 'Êtes-vous sûr de vouloir supprimer cette entrée ?',
                                              QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm_delete == QMessageBox.Yes:
            id_item = self.table_destinataire.item(row, 0)
            if id_item is not None:
                id_to_delete = id_item.text()
                try:
                    self.data_destination = [item for item in self.data_destination if item['Id'] != id_to_delete]
                    self.table_destinataire.removeRow(row)
                    for i in range(row, self.table_destinataire.rowCount()):
                        self.table_destinataire.item(i, 0).setText(str(i + 1))
                    self.save_destination_to_json()
                except Exception as e:
                    print("Erreur lors de la suppression:", str(e))
                    pass
                finally:
                    self.load_destination_from_json()

    def modifier_objet(self, no, nouveau_statut):
        try:
            with open('historique.json', "r") as f:
                data = json.load(f)

            for obj in data:
                if int(obj["No"]) == int(no):
                    obj["DateTime"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    obj["Status"] = nouveau_statut

            with open('historique.json', "w") as f:
                json.dump(data, f, indent=4)
            return True
        except Exception as e:
            print(f"Erreur de Modification : {str(e)}")
            return False

    def resend_row_from_table(self, row):
        confirm_delete = QMessageBox.question(self, 'Confirmation', 'Êtes-vous sûr de vouloir renvoyer cette mail ?',
                                              QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm_delete == QMessageBox.Yes:
            id_item = self.table_historique.item(row, 0)
            if id_item is not None:
                try:
                    id_to_resend = id_item.text()
                    new_status = True

                    self.modifier_objet(id_to_resend, new_status)
                    # if modif: print(f"Modification Faite sur N° {id_to_resend}")
                except Exception as e:
                    print(f"Erreur : {str(e)}")
                finally:
                    self.load_historique_from_json()

    def save_destination_to_json(self):
        data = []
        for row in range(self.table_destinataire.rowCount()):
            id_item = self.table_destinataire.item(row, 0)
            nom_item = self.table_destinataire.item(row, 1)
            email_item = self.table_destinataire.item(row, 2)

            # Get the status from the combo box and convert it to boolean
            status_combo = self.table_destinataire.cellWidget(row, 3)
            status = status_combo.currentText() == "Active"  # True if "Active", False otherwise

            if id_item is not None and nom_item is not None and email_item is not None:
                data.append({
                    'Id': id_item.text(),
                    'Nom': nom_item.text(),
                    'Email': email_item.text(),
                    'Status': status  # Add the converted boolean status
                })

        with open('destination.json', 'w', encoding='utf-8') as file:
            json.dump(data, file, indent=4)

        self.load_destination_from_json()

    def load_destination_from_json(self):
        try:
            with open('destination.json', 'r', encoding='utf-8') as file:
                data = json.load(file)

            # Effacer le contenu actuel du tableau
            self.table_destinataire.setRowCount(0)

            # Ajouter les données du fichier JSON au tableau
            for item in data:
                row_position = self.table_destinataire.rowCount()
                self.table_destinataire.insertRow(row_position)
                self.table_destinataire.setItem(row_position, 0, QTableWidgetItem(str(item['Id'])))
                self.table_destinataire.setItem(row_position, 1, QTableWidgetItem(item['Nom']))
                self.table_destinataire.setItem(row_position, 2, QTableWidgetItem(item['Email']))

                # Create a combo box with the appropriate status selected
                status_combo = QComboBox()
                status_combo.addItems(['Active', 'Inactive'])
                status_combo.setCurrentText(
                    "Active" if item['Status'] else "Inactive")  # Set based on loaded boolean status
                self.table_destinataire.setCellWidget(row_position, 3, status_combo)

                delete_button = QPushButton('X')
                delete_button.clicked.connect(lambda _, row=row_position: self.delete_row_from_table(row))
                self.table_destinataire.setCellWidget(row_position, 4, delete_button)

        except FileNotFoundError:
            print("Fichier Introuvable !")
            pass

    def load_historique_from_json(self):
        try:
            with open('historique.json', 'r', encoding='utf-8') as file:
                data = json.load(file)

            self.table_historique.setRowCount(0)

            for item in data:
                row_position = self.table_historique.rowCount()
                self.table_historique.insertRow(row_position)
                self.table_historique.setItem(row_position, 0, QTableWidgetItem(str(item['No'])))
                self.table_historique.setItem(row_position, 1, QTableWidgetItem(str(item['DateTime'])))
                self.table_historique.setItem(row_position, 2, QTableWidgetItem(str(item['Destinataire'])))
                self.table_historique.setItem(row_position, 3, QTableWidgetItem(str(item['Email'])))
                self.table_historique.setItem(row_position, 4,
                                              QTableWidgetItem("Success" if item['Status'] else "Echec"))
                if not item['Status']:
                    resend_button = QPushButton('^')
                    resend_button.clicked.connect(lambda _, row=row_position: self.resend_row_from_table(row))
                    self.table_historique.setCellWidget(row_position, 5, resend_button)
        except FileNotFoundError:
            print("Fichier Introuvable !")
            pass

    # affichage de la minuterie de compte à rebours en fonction du temps actuel et du temps cible
    def update_countdown(self):
        current_time = QDateTime.currentDateTime()
        if self.target_time is None:
            self.label.setText("Please set target time")
            return

        # Calculer la prochaine date et heure d'exécution en ajoutant un jour à la fois
        target_datetime = self.target_time
        while current_time >= target_datetime:
            target_datetime = target_datetime.addDays(1)

        diff = current_time.secsTo(target_datetime)

        if diff <= 0:
            self.iter_destination_json()
            self.recalculate_target_time()
            return

        days, remainder = divmod(diff, 86400)
        hours, remainder = divmod(remainder, 3600)
        minutes, seconds = divmod(remainder, 60)

        time_left = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        self.label.setText(time_left)

        if time_left == "00:00:00":
            return "Envoyé avec succès"
        return time_left; 
        

    def start_countdown(self):
        try:
            target_hour = int(self.target_hour_edit.text())
            target_minute = int(self.target_minute_edit.text())
            target_second = int(self.target_second_edit.text())

            self.target_time = QDateTime.currentDateTime()
            self.target_time.setTime(QTime(target_hour, target_minute, target_second))
            # self.target_day = self.day_combo.currentText()

            self.timer.start(1000)  # Update every second
            self.label.setText(self.get_time_remaining_text())
            self.startBtn.setEnabled(False)
            self.endBtn.setEnabled(True)
        except Exception as e:
            print(str(e))
            pass

    def end_timer(self):
        try:
            self.timer.stop()
            self.label.setText("Countdown Stopped")
            self.startBtn.setEnabled(True)
            self.endBtn.setEnabled(False)
        except Exception as e:
            print(str(e))
            pass

    def get_target_day_offset(self):
        return 1

    def recalculate_target_time(self):
        self.target_time = self.target_time.addDays(1)  # Ajoutez un jour au lieu de 7 jours
        self.label.setText(self.get_time_remaining_text())
        

    def get_time_remaining_text(self):
        return "Recalculating for Next day" if self.target_time is None else self.calculate_time_remaining()

    def calculate_time_remaining(self):
        current_time = QDateTime.currentDateTime()
        target_datetime = self.target_time.addDays(self.get_target_day_offset())

        diff_seconds = current_time.secsTo(target_datetime)

        if diff_seconds < 0:
            diff_days = current_time.daysTo(target_datetime)  # Calculate days from current to target
            if diff_days >= 0:
                diff_seconds += (diff_days * 86400)  # Adjust diff_seconds based on days difference
            else:
                diff_seconds = -diff_seconds  # Handle negative diff_seconds appropriately
        diff_days = diff_seconds // 86400  # Calculate remaining days

        diff_seconds %= 86400  # Get remaining seconds in the day

        hours, remainder = divmod(diff_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        time_left = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

        if time_left == "00:00:00":
            print('email envoyé')
        
        return time_left

    def iter_destination_json(self):
        # Load existing data from historique.json (or create an empty list if not found)
        try:
            with open('historique.json', 'r', encoding='utf-8') as file:
                historique_data = json.load(file)
        except FileNotFoundError:
            historique_data = []

        next_no = len(historique_data) + 1  # Start with the next available N°

        # Process destinations and create new entries
        save = []
        with open('destination.json', 'r', encoding='utf-8') as file:
            data = json.load(file)
        for item in data:
            if item['Status']:
                save.append({
                    'No': next_no,
                    'Destinataire': item['Nom'],
                    'Email': item['Email'],
                    'DateTime': datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    'Status': True
                })
                next_no += 1  # Increment N° for the next entry

        # Combine existing and new data (historique_data + save)
        combined_data = historique_data + save

        # Save the combined data to historique.json
        with open('historique.json', 'w', encoding='utf-8') as file:
            json.dump(combined_data, file, indent=4)

        self.load_historique_from_json()  # Reload data to update table


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = WinForm()
    win.show()
    sys.exit(app.exec_())
# data = getDateInSQLServer()
# if data is not None:
#     exportDataFrameEncaissement(data)
# else:
#     print("Data est vide !")
