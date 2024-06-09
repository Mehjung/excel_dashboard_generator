import xlwings as xw
import uuid

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
        self.slicer_cache = {} 
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
            TableDestination=sht_pivot.range('A11').api,
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
            print(pivot_table.PivotFields(field).Name)
            pivot_table.PivotFields(field).Orientation = xw.constants.PivotFieldOrientation.xlRowField
                
        print(f"Pivot-Tabelle '{pivot_table_name}' erstellt.")  

        

    def add_calculated_column(self, sheet_name, table_name, new_column_name, source_column_name, formula):
        sht = self.wb.sheets[sheet_name]
        tbl = sht.api.ListObjects(table_name)
        
        # Neue Spalte zur Tabelle hinzufügen
        tbl.ListColumns.Add().Name = new_column_name
        
        
        # Vollständige Formel für die neue Spalte erstellen
        full_formula = formula.replace("{source_column}", f"[{source_column_name}]")
        print (full_formula)
        # Formel auf die gesamte neue Spalte anwenden
        xw.Range(f'{table_name}[{new_column_name}]').formula = full_formula#'=IF([dienstart]="Spät",1,0)'
        
        

        print(f"Berechnete Spalte '{new_column_name}' mit Formel '{full_formula}' zu Tabelle '{table_name}' hinzugefügt.")

    def add_slicer(self, slicer_name, dashboard_sheet_name, pivot_sheet_name, pivot_table_name, top, left, width, height):
        sht_dashboard = self.wb.sheets[dashboard_sheet_name]
        sht_pivot = self.wb.sheets[pivot_sheet_name]
        pivot_table = sht_pivot.api.PivotTables(pivot_table_name)

        # Überprüfen, ob der Slicer-Cache bereits existiert
        slicer_cache = self.slicer_cache.get(slicer_name)

        # Wenn der Slicer-Cache nicht existiert, erstelle einen neuen und speichere ihn
        if slicer_cache is None:
            slicer_cache = self.wb.api.SlicerCaches.Add2(pivot_table, slicer_name)
            self.slicer_cache[slicer_name] = slicer_cache

        # Generiere eine eindeutige ID für den Slicer
        unique_slicer_name = f"{slicer_name}_{uuid.uuid4().hex}"

        # Füge den Slicer hinzu
        slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name=unique_slicer_name, Caption=slicer_name, Top=top, Left=left, Width=width, Height=height)
        slicer.Width = 100
        slicer.Height = 120

        # Freeze slicer position
        slicer.DisableMoveResizeUI = True

        print(f"Slicer '{unique_slicer_name}' hinzugefügt.")

    def create_dashboard(self, dashboard_sheet_name, pivot_sheet_name, pivot_table_name):
        # Dashboard-Blatt erstellen
        sht_dashboard = self.wb.sheets[dashboard_sheet_name]
        sht_dashboard.range('A1').value = "Analyse-Dashboard"
        sht_dashboard.range('A1').font.size = 24
        sht_dashboard.range('A1').font.bold = True

        # Hintergrundfarbe des Blattes ändern
        sht_dashboard.range('A1:Z100').color = (245, 245, 245)  # Helles Grau

        # Slicer hinzufügen und nebeneinander positionieren
        slicer_height = 200  # Höhe der Slicer anpassen
        self.add_slicer('betriebshof', dashboard_sheet_name, pivot_sheet_name, pivot_table_name, top=70, left=10, width=150, height=slicer_height)
        self.add_slicer('id', dashboard_sheet_name, pivot_sheet_name, pivot_table_name, top=70, left=170, width=150, height=slicer_height)
        self.add_slicer('langname', dashboard_sheet_name, pivot_sheet_name, pivot_table_name, top=200, left=10, width=150, height=slicer_height)

        # Balkendiagramm hinzufügen
        chart1 = sht_dashboard.charts.add(left=330, top=70, width=600, height=200)
        chart1.chart_type = 'bar_clustered'
        chart1.set_source_data(self.wb.sheets[pivot_sheet_name].range('A12').expand('table'))
        chart1.api[1].PlotArea.Format.Fill.ForeColor.RGB = 0xFFFFFF  # Weiß
        chart1.api[1].ChartArea.RoundedCorners = True
        chart1.api[1].PlotArea.Format.Line.Visible = True
        chart1.api[1].PlotArea.Format.Line.ForeColor.RGB = 0x000000  # Schwarz

        # Säulendiagramm hinzufügen
        chart2 = sht_dashboard.charts.add(left=330, top=300, width=600, height=200)
        chart2.chart_type = 'column_clustered'
        chart2.set_source_data(self.wb.sheets[pivot_sheet_name].range('A11').expand('table'))
        chart2.api[1].PlotArea.Format.Fill.ForeColor.RGB = 0xFFFFFF  # Weiß
        chart2.api[1].ChartArea.RoundedCorners = True
        chart2.api[1].PlotArea.Format.Line.Visible = True
        chart2.api[1].PlotArea.Format.Line.ForeColor.RGB = 0x000000  # Schwarz

        # Arbeitsmappe speichern
        self.wb.save()

        print(f"Dashboard '{dashboard_sheet_name}' erstellt.")



