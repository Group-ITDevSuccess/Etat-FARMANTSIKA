from openpyxl import Workbook


def exportDataFrame(data):
    wb = Workbook()
    ws = wb.active

    # Créer un dictionnaire pour stocker les données dans le format attendu
    formatted_data = {}
    for key, values in data.items():
        for value in values:
            intitule = value[0]
            amount = value[1]
            if intitule not in formatted_data:
                formatted_data[intitule] = {'day': 0, 'month': 0, 'year': 0}
            formatted_data[intitule][key] = amount

    # Écrire les en-têtes des colonnes
    ws.append(['Intitulé', 'day', 'month', 'year'])

    # Écrire les données dans le fichier Excel
    for intitule, amounts in formatted_data.items():
        ws.append([intitule, amounts['day'], amounts['month'], amounts['year']])

    # Sauvegarder le fichier Excel
    wb.save('data_output.xlsx')
