from data_repository import DataRepository
from excel_operations import ExcelOperations
import xlwings as xw


if __name__ == '__main__':
    repo_default = DataRepository()
    df = repo_default.get_data()


    with ExcelOperations() as excel_ops:
        #Füge ein neues Blatt hinzu
        source_sheet_name = 'Daten'
        sht = excel_ops.add_sheet(source_sheet_name)


        # Füge die Tabelle hinzu
        source_table_name = 'DatenTabelle'
        excel_ops.add_table(source_sheet_name, df, source_table_name)

        
        time_columns = ['dienstdauer', 'bezahlt', 'lenkzeit']
        excel_ops.format_time_columns(source_sheet_name, time_columns)

        # Ersetze Werte in der Spalte 'dienstart'
        excel_ops.replace_values_in_column(source_sheet_name, 'dienstart')

        pivot_table_name = 'Pivot-Tabelle'
        pivot_sheet_name = 'Pivot-Tabelle'
        pivot_sheet = excel_ops.add_sheet(pivot_sheet_name)

        row_fields = ["ID","Betriebshof","Dienstart"]
        data_fields = [('bezahlt_VZP', "='bezahlt' / ZEIT(7;48;0)", "0.00"),
                       ('bezahlt', xw.constants.ConsolidationFunction.xlSum,"[hh]:mm"),
                       ('lenkzeit', xw.constants.ConsolidationFunction.xlSum,"[hh]:mm"),
                       ('dienstdauer', xw.constants.ConsolidationFunction.xlAverage,"[hh]:mm"),
                       ('Betriebshof', xw.constants.ConsolidationFunction.xlCount,"0"),
                       ('Dienstart', xw.constants.ConsolidationFunction.xlCount,"0"),
                       ('Pausenregel', xw.constants.ConsolidationFunction.xlCount,"0")]
        excel_ops.create_pivot_table(pivot_table_name, source_sheet_name, source_table_name, pivot_sheet_name, row_fields, data_fields)