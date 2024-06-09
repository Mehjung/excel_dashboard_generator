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

        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='Spät',
        source_column_name='dienstart',
        formula='=IF({source_column}="Spät",1,0)'
        )

        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='Früh',
        source_column_name='dienstart',
        formula='=IF({source_column}="Früh",1,0)'
        )

        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='Tag',
        source_column_name='dienstart',
        formula='=IF({source_column}="Tag",1,0)'
        )
        
        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='1x30',
        source_column_name='Pausenregel',
        formula='=IF({source_column}="1x30",1,0)'
        )

        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='2x15',
        source_column_name='Pausenregel',
        formula='=IF({source_column}="2x15",1,0)'
        )

        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='1x45',
        source_column_name='Pausenregel',
        formula='=IF({source_column}="1x45",1,0)'
        )

        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='0x0',
        source_column_name='Pausenregel',
        formula='=IF({source_column}=0,1,0)'
        )

        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='<7:00',
        source_column_name='Dienstdauer',
        formula='=IF({source_column}<TIME(7,0,0),1,0)'
        )

        excel_ops.add_calculated_column(
        sheet_name=source_sheet_name,
        table_name=source_table_name,
        new_column_name='>10:00',
        source_column_name='Dienstdauer',
        formula='=IF({source_column}>TIME(10,0,0),1,0)'
        )

        pivot_table_name = 'Pivot-Tabelle'
        pivot_sheet_name = 'Pivot-Tabelle'
        pivot_sheet = excel_ops.add_sheet(pivot_sheet_name)

        row_fields = ["ID","Langname","Betriebshof"]
        data_fields = [('bezahlt_VZP', "='bezahlt' / ZEIT(7;48;0)", "0.00"),
                       ('bezahlt', xw.constants.ConsolidationFunction.xlSum,"[hh]:mm"),
                       ('lenkzeit', xw.constants.ConsolidationFunction.xlAverage,"[hh]:mm"),
                       ('dienstdauer', xw.constants.ConsolidationFunction.xlAverage,"[hh]:mm"),
                       ('Betriebshof', xw.constants.ConsolidationFunction.xlCount,"0"),
                       ('Spät', xw.constants.ConsolidationFunction.xlSum,"0"),
                       ('Früh', xw.constants.ConsolidationFunction.xlSum,"0"),
                       ('Tag', xw.constants.ConsolidationFunction.xlSum,"0"),
                       ('1x30', xw.constants.ConsolidationFunction.xlSum,"0"),
                       ('2x15', xw.constants.ConsolidationFunction.xlSum,"0"),
                       ('1x45', xw.constants.ConsolidationFunction.xlSum,"0"),
                       ('0x0', xw.constants.ConsolidationFunction.xlSum,"0"),
                       ('<7:00', xw.constants.ConsolidationFunction.xlSum,"0"),
                       ('>10:00', xw.constants.ConsolidationFunction.xlSum,"0")
                       ]
        
        excel_ops.create_pivot_table(pivot_table_name, source_sheet_name, source_table_name, pivot_sheet_name, row_fields, data_fields)
        
        # Füge den ersten Slicer hinzu
        excel_ops.add_slicer('betriebshof', pivot_sheet_name, pivot_sheet_name, pivot_table_name, 10, 10, 150, 60)

        # Füge den zweiten Slicer hinzu, mit ausreichendem Abstand zum ersten Slicer
        excel_ops.add_slicer('id', pivot_sheet_name, pivot_sheet_name, pivot_table_name, 10, 230, 150, 60)

        # Füge den zweiten Slicer hinzu, mit ausreichendem Abstand zum ersten Slicer
        excel_ops.add_slicer('langname', pivot_sheet_name, pivot_sheet_name, pivot_table_name, 10, 430, 150, 60)

        dash_sheet_name = 'Dashboard'
        dsht = excel_ops.add_sheet(dash_sheet_name)

        excel_ops.create_dashboard(dash_sheet_name, pivot_sheet,pivot_table_name)


        