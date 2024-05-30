import pandas as pd
import xlwings as xw

# Beispieldaten erstellen
data = {
    'Datum': pd.date_range(start='2024-01-01', periods=20, freq='D').tolist(),
    'Umlauf': ['Umlauf 1', 'Umlauf 2', 'Umlauf 3', 'Umlauf 4', 'Umlauf 5'] * 4,
    'Schicht': ['Früh', 'Spät', 'Nacht', 'Früh', 'Spät', 'Nacht', 'Früh', 'Spät', 'Nacht', 'Früh',
                'Früh', 'Spät', 'Nacht', 'Früh', 'Spät', 'Nacht', 'Früh', 'Spät', 'Nacht', 'Früh'],
    'Dienststelle': ['FF', 'HH', 'M', 'B', 'S', 'FF', 'HH', 'M', 'B', 'S', 'FF', 'HH', 'M', 'B', 'S', 'FF', 'HH', 'M', 'B', 'S'],
    'Schichtlänge': [8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9],
    'Pausenlänge': [30, 30, 30, 45, 45, 45, 15, 15, 15, 30, 30, 30, 45, 45, 45, 15, 15, 15, 30, 30],
    'Produktivitätsgrad': [0.95, 0.85, 0.80, 0.90, 0.88, 0.92, 0.94, 0.91, 0.89, 0.87, 0.93, 0.90, 0.88, 0.85, 0.83, 0.92, 0.91, 0.89, 0.87, 0.86]
}

df = pd.DataFrame(data)

# Excel-Datei erstellen
file_path = 'dynamic_excel_dashboard_with_slicers.xlsx'
wb = xw.Book()  # Neue Arbeitsmappe erstellen
sht_data = wb.sheets.add('Daten')
sht_data.range('A1').value = df

# Sicherstellen, dass die Daten als Tabelle formatiert sind
sht_data.range('A1').value = df.columns.tolist()  # Setzen der Spaltenüberschriften
sht_data.range('A2').value = df.values.tolist()  # Setzen der Daten
tbl = sht_data.api.ListObjects.Add(1, sht_data.range('A1').expand().api, None, 1)  # xlSrcRange = 1, xlYes = 1
tbl.Name = "DatenTabelle"

# Pivot-Table und Slicer erstellen
sht_pivot = wb.sheets.add('Pivot-Tabelle')
source_address = sht_data.range('A1').expand().address  # Holen Sie sich die Adresse der Tabelle
pivot_cache = wb.api.PivotCaches().Create(SourceType=1, SourceData=f"Daten!{source_address}")  # xlDatabase = 1
pivot_table = pivot_cache.CreatePivotTable(TableDestination=sht_pivot.range('A1').api, TableName="PivotTable1")

# Pivot-Tabelle konfigurieren
pivot_table.PivotFields("Umlauf").Orientation = 1  # xlRowField
pivot_table.PivotFields("Schicht").Orientation = 1  # xlRowField
pivot_table.PivotFields("Dienststelle").Orientation = 1  # xlRowField
pivot_table.PivotFields("Datum").Orientation = 2  # xlColumnField
pivot_table.AddDataField(pivot_table.PivotFields("Produktivitätsgrad"), "Durchschnittlicher Produktivitätsgrad", -4112)  # xlAverage

# Dashboard-Blatt erstellen
sht_dashboard = wb.sheets.add('Dashboard')
sht_dashboard.range('A1').value = "Produktivitäts-Dashboard"
sht_dashboard.range('A1').font.size = 24
sht_dashboard.range('A1').font.bold = True

# Balkendiagramm erstellen
chart = sht_dashboard.charts.add(250, 50)
chart.chart_type = 'column_clustered'
chart.set_source_data(sht_pivot.range('A1').expand())
chart.name = "Produktivitätsgrad nach Umlauf"
chart.api.HasTitle = True
chart.api.ChartTitle.Text = "Produktivitätsgrad nach Umlauf"

# Slicer hinzufügen
sht_pivot.range('A1').select()
slicer_cache = wb.api.SlicerCaches.Add(pivot_table, pivot_table.PivotFields("Umlauf"))
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, "", "Umlauf", "Umlauf", 30, 30, 100, 100)
slicer_cache = wb.api.SlicerCaches.Add(pivot_table, pivot_table.PivotFields("Schicht"))
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, "", "Schicht", "Schicht", 30, 150, 100, 100)
slicer_cache = wb.api.SlicerCaches.Add(pivot_table, pivot_table.PivotFields("Dienststelle"))
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, "", "Dienststelle", "Dienststelle", 30, 270, 100, 100)

# Datei speichern
wb.save(file_path)
wb.close()

print(f'Dynamisches Excel Dashboard mit Slicern erstellt und gespeichert in {file_path}')
