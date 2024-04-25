import pyodbc

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


def getDateInSQLServer():
    connexion = connexionSQlServer(server="Srv-sagei7-004", base="FARMANTSIKA2020")
    if connexion:
        try:
            finds = extract_values_in_json('SQL')
            # lists = finds['lists']
            data = {}
            for key, value in finds.items():
                if key not in ['lists', 'details']:
                    with connexion.cursor() as cursor:
                        query = str(value).replace("{today}", get_today())
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


