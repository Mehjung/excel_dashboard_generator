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
file_path = 'professional_excel_dashboard.xlsx'
wb = xw.Book()  # Neue Arbeitsmappe erstellen
sht_data = wb.sheets.add('Daten')
sht_data.range('A1').value = df

# Sicherstellen, dass die Daten als Tabelle formatiert sind
tbl = sht_data.api.ListObjects.Add(1, sht_data.range('A1').expand().api, None, 1)  # xlSrcRange = 1, xlYes = 1
tbl.Name = "DatenTabelle"

# Pivot-Table und Slicer erstellen
sht_pivot = wb.sheets.add('Pivot-Tabelle')
source_address = tbl.Range.Address  # Korrigiert: Verwendung der Eigenschaft Address
pivot_cache = wb.api.PivotCaches().Create(SourceType=xw.constants.PivotTableSourceType.xlDatabase, SourceData=f"Daten!{source_address}")
pivot_table = pivot_cache.CreatePivotTable(TableDestination=sht_pivot.range('A3').api, TableName="PivotTable1")

# Pivot-Tabelle konfigurieren
pivot_table.PivotFields("Umlauf").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("Schicht").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("Dienststelle").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("Datum").Orientation = xw.constants.PivotFieldOrientation.xlColumnField
pivot_table.AddDataField(pivot_table.PivotFields("Produktivitätsgrad"), "Durchschnittlicher Produktivitätsgrad", xw.constants.ConsolidationFunction.xlAverage)

# Dashboard-Blatt erstellen
sht_dashboard = wb.sheets.add('Dashboard')
sht_dashboard.range('A1').value = "Produktivitäts-Dashboard"
sht_dashboard.range('A1').font.size = 24
sht_dashboard.range('A1').font.bold = True

# Diagramme erstellen
chart1 = sht_dashboard.charts.add(left=10, top=50, width=300, height=200)
chart1.chart_type = 'pie'
chart1.set_source_data(sht_pivot.range('A3').expand('table'))

chart2 = sht_dashboard.charts.add(left=320, top=50, width=300, height=200)
chart2.chart_type = 'column_clustered'
chart2.set_source_data(sht_pivot.range('A3').expand('table'))

chart3 = sht_dashboard.charts.add(left=10, top=260, width=300, height=200)
chart3.chart_type = 'bar_clustered'
chart3.set_source_data(sht_pivot.range('A3').expand('table'))

chart4 = sht_dashboard.charts.add(left=320, top=260, width=300, height=200)
chart4.chart_type = 'line'
chart4.set_source_data(sht_pivot.range('A3').expand('table'))

# Slicer hinzufügen
slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Umlauf")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Umlauf", Caption="Umlauf", Top=470, Left=10, Width=100, Height=100)
slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Schicht")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Schicht", Caption="Schicht", Top=470, Left=120, Width=100, Height=100)
slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Dienststelle")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Dienststelle", Caption="Dienststelle", Top=470, Left=230, Width=100, Height=100)

# Datei speichern
wb.save(file_path)
wb.close()

print(f'Professionelles Excel Dashboard mit Slicern erstellt und gespeichert in {file_path}')
