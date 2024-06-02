import xlwings as xw

class ExcelOperations:
    def __init__(self, config=None):
        if config is None:
            config = {
                'file_path': 'professional_excel_dashboard.xlsx',
                'visible': False,
                'alerts': False,
                'screen_updating': False
            }
        
        self.config = config
        self.file_path = self.config['file_path']
        self.app = xw.App(visible=self.config['visible'])
        self.app.display_alerts = self.config['alerts']
        self.app.screen_updating = self.config['screen_updating']
        self.wb = self.app.books.add()  # Erstellt eine neue Arbeitsmappe
        print(f"Excel-Datei erstellt: {self.file_path}")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            self.wb.save(self.file_path)
            self.wb.close()
            self.app.quit()
            print(f"Excel-Datei gespeichert und geschlossen: {self.file_path}")
        except Exception as e:
            print(f"Fehler beim Schließen der Excel-Anwendung: {e}")

    def add_sheet(self, sheet_name):
        # Überprüfen, ob das Blatt bereits existiert und löschen
        if sheet_name in [sht.name for sht in self.wb.sheets]:
            self.wb.sheets[sheet_name].delete()
        # Neues Blatt hinzufügen
        sht = self.wb.sheets.add(sheet_name)
        print(f"Blatt '{sheet_name}' hinzugefügt.")
        return sht

    def add_table(self, sheet_name, dataframe, table_name):
        sht = self.wb.sheets[sheet_name]
        sht.range('A1').value = dataframe

        # Sicherstellen, dass die Daten als Tabelle formatiert sind
        tbl = sht.api.ListObjects.Add(1, sht.range('A1').expand().api, None, 1)  # xlSrcRange = 1, xlYes = 1
        tbl.Name = table_name
        
        print(f"Tabelle '{table_name}' auf Blatt '{sheet_name}' hinzugefügt.")
        return tbl

    def format_time_columns(self, sheet_name, time_columns):
        sht = self.wb.sheets[sheet_name]
        
        # Setze das Zellenformat für Timedelta-Spalten
        for col in time_columns:
            col_range = sht.range('1:1').value
            if col in col_range:
                col_index = col_range.index(col) + 1
                col_letter = xw.utils.col_name(col_index)
                print(f"Formatierung der Spalte '{col}' mit Index {col_index} ({col_letter}) auf '[hh]:mm' gesetzt.")
                sht.range(f'{col_letter}2:{col_letter}{sht.range("A1").expand("down").rows.count}').number_format = '[hh]:mm'

    def replace_values_in_column(self, sheet_name, column_name):
        sht = self.wb.sheets[sheet_name]
        
        # Überprüfen, ob die Spalte im Blatt vorhanden ist
        col_range = sht.range('1:1').value
        if column_name in col_range:
            col_index = col_range.index(column_name) + 1
            col_letter = xw.utils.col_name(col_index)
            column_range = sht.range(f'{col_letter}2:{col_letter}{sht.range("A1").expand("down").rows.count}')
            values = column_range.value
            
            # Ersetze die Werte in der Spalte
            for i, cell in enumerate(column_range):
                value = values[i]
                if isinstance(value, str):
                    if value.endswith('S'):
                        cell.value = 'Spät'
                    elif value.endswith('F'):
                        cell.value = 'Früh'
                    elif value.endswith('T'):
                        cell.value = 'Tag'
            
            print(f"Werte in der Spalte '{column_name}' ersetzt.")

    def create_pivot_table(self, pivot_table_name, source_sheet_name, source_table_name, pivot_sheet_name, row_fields, data_fields):
        sht_pivot = self.wb.sheets[pivot_sheet_name]
        sht_source = self.wb.sheets[source_sheet_name]
        tbl = sht_source.api.ListObjects(source_table_name)
        source_address = tbl.Range.Address

        pivot_cache = self.wb.api.PivotCaches().Create(
            SourceType=xw.constants.PivotTableSourceType.xlDatabase,
            SourceData=f"{source_sheet_name}!{source_address}"
        )
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=sht_pivot.range('A3').api,
            TableName=pivot_table_name
        )

        for field, func, format in data_fields:
            if isinstance(func, str) and func.startswith("="):
                # Add a custom calculated field
                calculated_field_name = f"{field}_custom"
                pivot_table.CalculatedFields().Add(Name=calculated_field_name, Formula=func)
                data_field = pivot_table.AddDataField(pivot_table.PivotFields(calculated_field_name))
                data_field.Function = xw.constants.ConsolidationFunction.xlSum  # Default to Sum for custom
                data_field.NumberFormat = format
            else:
                data_field = pivot_table.AddDataField(pivot_table.PivotFields(field))
                data_field.Function = func
                data_field.NumberFormat = format

            # Set the row fields for the pivot table
        for field in row_fields:
            pivot_table.PivotFields(field).Orientation = xw.constants.PivotFieldOrientation.xlRowField
        print(f"Pivot-Tabelle '{pivot_table_name}' erstellt.")     
        