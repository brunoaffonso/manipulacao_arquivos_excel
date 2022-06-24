from datetime import date
from openpyxl.chart import Reference
from openpyxl.styles import Font, PatternFill, Alignment

try:
    # stock = input('Código da Ação: ').upper()
    from classes import StocksReader, SpreadsheetManager, ChartSeriesProperties

    stock = 'BIDI4a'

    stocks_reader = StocksReader(path='./dados/')
    stocks_reader.file_process(stock)

    manager = SpreadsheetManager()
    data_spreadsheet = manager.add_spreadsheet('Dados')

    manager.add_line(['DATA', 'COTAÇÃO', 'BANDA INFERIOR', 'BANDA SUPERIOR'])

    index = 2

    for line in stocks_reader.data:
        # Data
        year_month_day = line[0].split(' ')[0]
        date_ = date(
            year=int(year_month_day.split('-')[0]),
            month=int(year_month_day.split('-')[1]),
            day=int(year_month_day.split('-')[2])
        )

        # Price
        price = float(line[1])

        bb_higher_formule = f'=AVERAGE(B{index}:B{index + 19}) - 2*STDEV(B{index}:B{index + 19})'
        bb_bottom_formule = f'=AVERAGE(B{index}:B{index + 19}) + 2*STDEV(B{index}:B{index + 19})'

        # Update cells from active spreadsheet
        manager.update_cell(cell=f'A{index}', data=date_)
        manager.update_cell(cell=f'B{index}', data=price)
        manager.update_cell(cell=f'C{index}', data=bb_bottom_formule)
        manager.update_cell(cell=f'D{index}', data=bb_higher_formule)

        index += 1

    manager.add_spreadsheet('Gráfico')

    # Mergin header cells
    manager.merge_spreadsheet_cells(start_cell='A1', end_cell='T2')

    manager.apply_style(
        cell='A1',
        styles=[
            ('font', Font(name='Calibri', b=True, sz=18, color='FFFFFF')),
            ('fill', PatternFill('solid', fgColor='07838F')),
            ('alignment', Alignment(vertical='center', horizontal='center')),
        ]
    )

    manager.update_cell('A1', 'Histórico de Cotações')

    price_references = Reference(data_spreadsheet, min_col=2, min_row=2, max_col=4, max_row=index)
    date_references = Reference(data_spreadsheet, min_col=1, min_row=2, max_col=1, max_row=index)

    manager.add_line_chart(
        cell='A3',
        width=33.87,
        height=14.82,
        title=f'Cotações - {stock}',
        x_axis_title='Data da Cotação',
        y_axis_title='Valor da Cotação',
        x_axis_reference=price_references,
        y_axis_reference=date_references,
        chart_properties=[
            ChartSeriesProperties(width=0, solid_fill_color='0A55AB'),
            ChartSeriesProperties(width=0, solid_fill_color='A61508'),
            ChartSeriesProperties(width=0, solid_fill_color='12A154'),
        ]
    )

    manager.merge_spreadsheet_cells(start_cell='I32', end_cell='L35')
    manager.add_spreadsheet_image(cell='I32', image_path='./recursos/python.png')

    manager.save_file('./saida/planilha_refactor.xlsx')

except AttributeError:
    print('Atributo não existe!')

except ValueError:
    print('Formato de dados incorreto. Necessita verificação!')

except FileNotFoundError:
    print('Arquivo não encontrado!')

except Exception as excecao:
    print(f'Ocorreu um erro durante a execução do programa. Erro: {excecao}')
