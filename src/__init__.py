import datetime

from openpyxl import load_workbook

from graphviz import Digraph


def main():
    workSheet = load_workbook(
        '/Users/ninroot/OneDrive - EPITA/ING2/URSI/FNAC DARTY/Urbanisation shared/Urbanisation shared all/Matrice des flux.xlsx'
        , data_only=True)
    # print(workSheet.get_sheet_names())
    matrix = workSheet["Matrice des flux"]

    update = datetime.datetime.now()
    title = 'Matrice des flux (' + update.strftime('%Y/%m/%d %H:%M:%S') + ')'
    dot = Digraph(comment=title)
    dot.engine = 'circo'
    # dot.engine = 'sfdp'

    dot.node('Titre', title)

    dot.node('Backoffice', 'Backoffice')
    dot.node('Caisse', 'Caisse')
    dot.node('Fidélité', 'Fidélité')
    dot.node('Référentiel', 'Référentiel')
    dot.node('Monétique', 'Monétique')
    dot.node('eCommerce', 'eCommerce')
    dot.node('Entrepôt', 'Entrepôt')
    dot.node('BI', 'BI')
    dot.node('Réappro', 'Réappro')
    dot.node('Banque', 'Banque')

    for r in range(2, 42):
        row = matrix['A' + str(r) + ':' + 'M' + str(r)]
        if None in row[0]:
            print("None in row %s", str(r))
            continue
        # print(str(r) + ": " + row[0][0].value)
        # print(str(r) + ": " + row[0][1].value)
        dot.edge(row[0][0].value, row[0][1].value, label=row[0][6].value)
        # print(row_range)
        # for col in range(1, 13): # M
        #    print(matrix.cell(column=col, row=row).value)

    # print(dot.source)
    dot.render('matrix.gv', view=True)


if __name__ == "__main__":
    main()
