import pandas as pd
import xlwings as xw
from mockdata import data  # Import the data from data.py

df = pd.DataFrame(data)

# Excel-Datei erstellen
file_path = 'professional_excel_dashboard.xlsx'
wb = xw.Book()  # Neue Arbeitsmappe erstellen
xw.apps.active.api.Calculation = xw.constants.Calculation.xlCalculationAutomatic

sht_data = wb.sheets.add('Daten')
sht_data.range('A1').value = df

# Sicherstellen, dass die Daten als Tabelle formatiert sind
tbl = sht_data.api.ListObjects.Add(1, sht_data.range('A1').expand().api, None, 1)  # xlSrcRange = 1, xlYes = 1
tbl.Name = "DatenTabelle"


# Dynamisch nach der Spalte "Dienstart" suchen
header = sht_data.range('A1').expand('right').value
dienstart_col_index = header.index('Dienstart') + 1
dienstart_col_letter = xw.utils.col_name(dienstart_col_index)
print(f'Dienstart-Spalte gefunden: {dienstart_col_letter}')
#pritn col letter

sht_data.range(f'{xw.utils.col_name(len(header) + 1)}1').value = "Gemappte Dienstart"


# Formel für 'Gemappte Dienstart' setzen
last_row = sht_data.range('A1').end('down').row
sht_data.range(f'{xw.utils.col_name(len(header) + 1)}2:{xw.utils.col_name(len(header) + 1)}{last_row}').formula = f'=WENN(RECHTS({dienstart_col_letter}2,1)="T","Tag",WENN(RECHTS({dienstart_col_letter}2,1)="F","Früh",WENN(RECHTS({dienstart_col_letter}2, 1)="S","Spät","Andere")))'






# Aktualisieren Sie die Spaltenüberschriften
sht_data.range("G1").value = "Gemappte Dienstart"
sht_data.range("H1").value = "Dienstdauer Kategorie"

# Pivot-Table und Slicer erstellen
sht_pivot = wb.sheets.add('Pivot-Tabelle')
source_address = tbl.Range.Address  # Korrigiert: Verwendung der Eigenschaft Address
pivot_cache = wb.api.PivotCaches().Create(SourceType=xw.constants.PivotTableSourceType.xlDatabase, SourceData=f"Daten!{source_address}")
pivot_table = pivot_cache.CreatePivotTable(TableDestination=sht_pivot.range('A3').api, TableName="PivotTable1")

# Pivot-Tabelle konfigurieren
pivot_table.PivotFields("Betriebshof").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("Gemappte Dienstart").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("Pausenregel").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("Dienstdauer Kategorie").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("bezahlt").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("Lenkzeit").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("ID").Orientation = xw.constants.PivotFieldOrientation.xlRowField
pivot_table.PivotFields("Datum").Orientation = xw.constants.PivotFieldOrientation.xlColumnField
pivot_table.AddDataField(pivot_table.PivotFields("Kosten"), "Durchschnittliche Kosten", xw.constants.ConsolidationFunction.xlAverage)

# Dashboard-Blatt erstellen
sht_dashboard = wb.sheets.add('Dashboard')
sht_dashboard.range('A1').value = "Kosten-Dashboard"
sht_dashboard.range('A1').font.size = 24
sht_dashboard.range('A1').font.bold = True

# Hintergrundfarbe des Blattes ändern
sht_dashboard.range('A1:Z100').color = (255, 255, 255)  # Weiß

# Diagramme erstellen und abrunden
chart1 = sht_dashboard.charts.add(left=20, top=60, width=300, height=200)
chart1.chart_type = 'pie'
chart1.set_source_data(sht_pivot.range('A3').expand('table'))
chart1.api[1].PlotArea.Format.Fill.ForeColor.RGB = 0xFFFFFF  # Weiß
chart1.api[1].ChartArea.RoundedCorners = True
chart1.api[1].ChartTitle.Text = "Kosten"
chart1.api[1].ChartTitle.Font.Size = 14
chart1.api[1].ChartTitle.Font.Color = 0x000000  # Schwarz
chart1.api[1].HasLegend = True

chart2 = sht_dashboard.charts.add(left=330, top=60, width=300, height=200)
chart2.chart_type = 'column_clustered'
chart2.set_source_data(sht_pivot.range('A3').expand('table'))
chart2.api[1].PlotArea.Format.Fill.ForeColor.RGB = 0xFFFFFF  # Weiß
chart2.api[1].ChartArea.RoundedCorners = True

chart3 = sht_dashboard.charts.add(left=20, top=270, width=300, height=200)
chart3.chart_type = 'bar_clustered'
chart3.set_source_data(sht_pivot.range('A3').expand('table'))
chart3.api[1].PlotArea.Format.Fill.ForeColor.RGB = 0xFFFFFF  # Weiß
chart3.api[1].ChartArea.RoundedCorners = True

chart4 = sht_dashboard.charts.add(left=330, top=270, width=300, height=200)
chart4.chart_type = 'line'
chart4.set_source_data(sht_pivot.range('A3').expand('table'))
chart4.api[1].PlotArea.Format.Fill.ForeColor.RGB = 0xFFFFFF  # Weiß
chart4.api[1].ChartArea.RoundedCorners = True

# Slicer hinzufügen
slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Betriebshof")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Betriebshof", Caption="Betriebshof", Top=480, Left=20, Width=100, Height=100)

slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Gemappte Dienstart")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Gemappte_Dienstart", Caption="Gemappte Dienstart", Top=480, Left=130, Width=100, Height=100)

slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Pausenregel")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Pausenregel", Caption="Pausenregel", Top=480, Left=240, Width=100, Height=100)

slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Dienstdauer Kategorie")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Dienstdauer_Kategorie", Caption="Dienstdauer Kategorie", Top=480, Left=350, Width=100, Height=100)

slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "bezahlt")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="bezahlt", Caption="bezahlt", Top=480, Left=460, Width=100, Height=100)

slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Lenkzeit")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Lenkzeit", Caption="Lenkzeit", Top=480, Left=570, Width=100, Height=100)

slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "ID")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="ID", Caption="ID", Top=480, Left=680, Width=100, Height=100)

slicer_cache = wb.api.SlicerCaches.Add2(sht_pivot.api.PivotTables("PivotTable1"), "Thema")
slicer = slicer_cache.Slicers.Add(sht_dashboard.api, Name="Thema", Caption="Thema", Top=480, Left=790, Width=100, Height=100)

# Datei speichern
wb.save(file_path)
wb.close()

print(f'Professionelles Excel Dashboard mit Slicern erstellt und gespeichert in {file_path}')
