import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

# Beispieldaten erstellen
data = {
    'Datum': ['2024-01-01', '2024-01-01', '2024-02-01', '2024-02-01', '2024-03-01', '2024-03-01'],
    'Produkt': ['Produkt A', 'Produkt B', 'Produkt A', 'Produkt B', 'Produkt A', 'Produkt B'],
    'Region': ['Nord', 'Süd', 'Nord', 'Süd', 'Nord', 'Süd'],
    'Umsatz': [100, 200, 150, 250, 200, 300]
}

df = pd.DataFrame(data)

# Excel-Datei erstellen
file_path = 'dynamic_excel_dashboard_with_filter.xlsx'
df.to_excel(file_path, index=False, sheet_name='Daten')

# Arbeitsmappe laden
wb = load_workbook(file_path)
ws = wb.active

# AutoFilter hinzufügen
ws.auto_filter.ref = ws.dimensions

# Pivot-Tabelle erstellen
pivot_data = pd.pivot_table(df, values='Umsatz', index=['Produkt'], columns=['Region'], aggfunc='sum', fill_value=0)
pivot_ws = wb.create_sheet(title='Pivot-Tabelle')

for r in dataframe_to_rows(pivot_data, index=True, header=True):
    pivot_ws.append(r)

# Formatierungen und Stile anwenden
header_font = Font(name='Arial', bold=True, color='FFFFFF', size=14)
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
cell_alignment = Alignment(horizontal='center', vertical='center')

for cell in ws['1:1']:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = cell_alignment

# Zusätzliche Formatierung für Pivot-Tabelle
pivot_ws['A1'].font = Font(bold=True)
pivot_ws['B1'].font = Font(bold=True)
pivot_ws['C1'].font = Font(bold=True)
for cell in pivot_ws['A:A']:
    cell.font = Font(bold=True)
pivot_ws['A1'].fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
pivot_ws['B1'].fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
pivot_ws['C1'].fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')

# Bedingte Formatierung hinzufügen
red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
pivot_ws.conditional_formatting.add('B2:C10', CellIsRule(operator='lessThan', formula=['150'], fill=red_fill))
pivot_ws.conditional_formatting.add('B2:C10', CellIsRule(operator='greaterThanOrEqual', formula=['150'], fill=green_fill))

# Dynamische Diagramme in Excel erstellen

# Balkendiagramm erstellen
bar_chart = BarChart()
data = Reference(ws, min_col=4, min_row=1, max_col=4, max_row=len(df) + 1)
categories = Reference(ws, min_col=1, min_row=2, max_row=len(df) + 1)
bar_chart.add_data(data, titles_from_data=True)
bar_chart.set_categories(categories)
bar_chart.title = "Umsatz nach Monat"
ws.add_chart(bar_chart, "E10")

# Liniendiagramm erstellen
line_chart = LineChart()
data = Reference(ws, min_col=4, min_row=1, max_col=4, max_row=len(df) + 1)
line_chart.add_data(data, titles_from_data=True)
line_chart.set_categories(categories)
line_chart.title = "Umsatzentwicklung nach Produkt"
ws.add_chart(line_chart, "E30")

# Kreisdiagramm erstellen
pie_chart = PieChart()
data = Reference(ws, min_col=4, min_row=1, max_col=4, max_row=len(df) + 1)
pie_chart.add_data(data, titles_from_data=True)
categories = Reference(ws, min_col=2, min_row=2, max_row=len(df) + 1)
pie_chart.set_categories(categories)
pie_chart.title = "Umsatzverteilung nach Produkt"
ws.add_chart(pie_chart, "E50")

# Arbeitsmappe speichern
wb.save(file_path)

print(f'Dynamisches Excel Dashboard mit Filter erstellt und gespeichert in {file_path}')
