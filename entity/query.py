import os
from copy import copy

import pyodbc
import pandas as pd
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import NamedStyle
from openpyxl.workbook import Workbook

from entity.export import exportDataFrameEncaissement, exportDataFrameDetails, merge_excel_files
from utils import write_log, extract_values_in_json, get_today


def connexionSQlServer(server, base):
    conn = None
    value = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={base};UID=reader;PWD=m1234"
    try:
        conn = pyodbc.connect(value)
    except pyodbc.Error as e:
        # Si une erreur se produit lors de la connexion, ajoutez le serveur à la liste des bases sans accès
        print("===============================")
        print(f"Erreur de Connexion pour {server} à la base {base} ")
        print("===============================")
        write_log(f"Erreur de connexion : {str(e)}")
    except Exception as e:
        write_log(f"Erreur Exception : {str(e)}")
    return conn


def getDataLink(connexion):
    if connexion:
        try:
            finds = extract_values_in_json('SQL')
            dates = get_today()
            resumer = {}
            details = None
            for key, value in finds.items():
                if key not in ['details']:
                    with connexion.cursor() as cursor:
                        query = str(value).replace(
                            "{today}", dates.strftime('%m/%d/%Y')
                        ).replace("{year}", dates.strftime('%Y')
                                  ).replace('{month}', dates.strftime('%m'))
                        cursor.execute(query)
                        rows = cursor.fetchall()
                        tuples = []
                        if rows:
                            tuples = [tuple(row) for row in rows]
                        resumer[key] = tuples

                elif key in ['details']:
                    with connexion.cursor() as cursor:
                        query = str(value).replace("{today}", dates.strftime('%m/%d/%Y'))
                        cursor.execute(query)
                        rows = cursor.fetchall()
                        rows = [tuple(row) for row in rows]
                        if all(isinstance(row, tuple) for row in rows):
                            details = pd.DataFrame(rows,
                                                   columns=["N° PIECE", "REF", "DESIGNATION", "QTE", "PRIX UNITAIRE",
                                                            "MONTANT TTC",
                                                            "MODE"])
                        else:
                            print("Pas un tuple")
            excel_1 = exportDataFrameEncaissement(resumer)
            excel_2 = exportDataFrameDetails(details)

            storage_dir = f'storage/{get_today().strftime("%d-%m-%Y")}'
            if not os.path.exists(storage_dir):
                os.makedirs(storage_dir)

            wb1 = load_workbook(excel_1)
            wb2 = load_workbook(excel_2)

            for sheet in wb2.sheetnames:
                ws = wb2[sheet]
                if sheet not in wb1.sheetnames:
                    wb1.create_sheet(title=sheet)

                # Copier les styles de l'Excel_2 vers l'Excel_1
                for row in ws.iter_rows():
                    for cell in row:
                        dest_cell = wb1[sheet][cell.coordinate]
                        dest_cell.font = copy(cell.font)
                        dest_cell.fill = copy(cell.fill)
                        dest_cell.border = copy(cell.border)
                        dest_cell.alignment = copy(cell.alignment)
                        dest_cell.number_format = copy(cell.number_format)
                        dest_cell.protection = copy(cell.protection)

                # Copier les données de l'Excel_2 vers l'Excel_1 dans les mêmes cellules
                for row in ws.iter_rows():
                    for cell in row:
                        dest_cell = wb1[sheet][cell.coordinate]
                        dest_cell.value = cell.value

            # Espacer les colonnes en fonction du contenu
            for sheet in wb1.sheetnames:
                ws = wb1[sheet]
                for column_cells in ws.columns:
                    max_length = 0
                    for cell in column_cells:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2) * 1.2
                    ws.column_dimensions[column_cells[0].column_letter].width = adjusted_width

            # Supprimer la première feuille vide créée automatiquement dans wb1
            if 'Sheet' in wb1.sheetnames:
                wb1.remove(wb1['Sheet'])

            # Enregistrer le classeur Excel fusionné dans le dossier "storage"
            merged_file_path = os.path.join(storage_dir, f'ETAT FARMANTSIKA_{get_today().strftime("%d-%m-%Y")}.xlsx')
            wb1.save(merged_file_path)

            print(f"Fichier fusionné enregistré à : {merged_file_path}")

            return merged_file_path
        except Exception as e:
            write_log(f"Erreur Exception : {str(e)}")
        finally:
            connexion.close()

        return None


def getDetailInSQLServer(connexion):
    data = {}

    if connexion:
        try:
            finds = extract_values_in_json('SQL')
            dates = get_today()
            for key, value in finds.items():
                if key in ['details']:
                    with connexion.cursor() as cursor:
                        query = str(value).replace(
                            "{today}", dates.strftime('%m/%d/%Y')
                        ).replace("{year}", dates.strftime('%Y')
                                  ).replace('{month}', dates.strftime('%m'))
                        print(f"Query : {query}")
                        cursor.execute(query)
                        rows = cursor.fetchall()
                        tuples = []
                        if rows:
                            tuples = [tuple(row) for row in rows]
                        data[key] = tuples

            print(f"DATA: {data}")
        except Exception as e:
            write_log(f"Erreur Exception : {str(e)}")
        finally:
            connexion.close()

        return data
