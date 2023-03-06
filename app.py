import openpyxl
from product_to_barcode import product_to_barcode

planilha_vendas = "Resumo de Vendas por clientes Nov Dez Jan TRUST HAIR SOLUTIONS.xlsx"
read_sells = openpyxl.load_workbook(planilha_vendas)
read_sells_sheet = read_sells["relFaturamentoPedidosResumidoPr"]

products_list = []
for row in range(2, read_sells_sheet.max_row + 1):
    for column in "E":
        cell_name = "{}{}".format(column, row)
        product = read_sells_sheet[cell_name].value
        if product != None and product != "Sub-Total" and product != "Total Geral" and product != "" and "L´ARREÉ" not in product and "L´ARRÉ" not in product and "CX Papel Ecologico" not in product and "BlueMoon" not in product:
            for column in "H":
                h_cell_name = "{}{}".format(column, row)
                read_sells_sheet[h_cell_name].value = product_to_barcode(product)


read_sells.save("Nova Planilha Vendas Com Código de Barras.xlsx")