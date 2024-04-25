from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font

from utils import get_today


def exportDataFrameEncaissement(data):
    wb = Workbook()
    ws = wb.active
    today = get_today()
    # Renommer la feuille de calcul
    ws.title = "ENCAISSEMENT"

    # Créer un dictionnaire pour stocker les données dans le format attendu
    formatted_data = {}
    for key, values in data.items():
        for value in values:
            intitule = value[0]
            amount = value[1]
            if intitule not in formatted_data:
                formatted_data[intitule] = {'day': 0, 'month': 0, 'year': 0}
            formatted_data[intitule][key] = amount

    # Écrire les en-têtes des colonnes avec mise en forme
    header_row = ['Intitulé', f'{today.strftime('%d/%m/%Y')}', 'Cumul Mois', 'Cumul Année']
    ws.append(header_row)
    for cell in ws[1]:
        cell.font = Font(bold=True)  # Rendre l'entête en gras

    # Écrire les données dans le fichier Excel avec mise en forme
    for intitule, amounts in formatted_data.items():
        ws.append([intitule, amounts['day'], amounts['month'], amounts['year']])

    # Ajouter des bordures aux cellules
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                                 left=Side(style='thin'), right=Side(style='thin'))

    # Colorier les cellules
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=ws.max_column):
        for cell in row:
            cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    # Sauvegarder le fichier Excel
    wb.save('data_output.xlsx')