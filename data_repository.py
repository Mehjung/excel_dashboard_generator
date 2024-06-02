import pandas as pd
import xlwings as xw

class DataRepository:
    def __init__(self, file_path='data.xlsx', sheet_name='Consolidate'):
        self.config = {
            'file_path': file_path,
            'sheet_name': sheet_name
        }
        self.df = self.load_and_format_data()

    def load_and_format_data(self):
        try:
            file_path = self.config['file_path']
            sheet_name = self.config['sheet_name']
            
            print(f"Lade Datei: {file_path}")
            app = xw.App(visible=False)
            app.display_alerts = False  # Deaktiviert Pop-ups
            app.screen_updating = False  # Deaktiviert Bildschirmaktualisierungen
            
            wb = app.books.open(file_path, update_links=False)  # Öffnet die Datei ohne Links zu aktualisieren
            sht = wb.sheets[sheet_name]
            print(f"Lade Sheet: {sheet_name}")
            
            data = sht.range('A1').expand().value
            
            # Laden der Daten in einen DataFrame
            df = pd.DataFrame(data[1:], columns=data[0])
            wb.close()
            app.quit()
            
            print(f"Daten geladen: erfolgreich ({df.shape[0]} Zeilen, {df.shape[1]} Spalten)")
            
            # Normalisieren der Spaltennamen
            df.columns = [col.strip().lower() for col in df.columns]
            
            return df
        except Exception as e:
            print(f"Fehler beim Laden der Daten: {e}")
            return pd.DataFrame()  # Leerer DataFrame bei Fehler

    def get_data(self):
        return self.df