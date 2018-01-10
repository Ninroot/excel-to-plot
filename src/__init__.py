import datetime

from openpyxl import load_workbook

from graphviz import Digraph


def main():
    workSheet = load_workbook('/Users/ninroot/OneDrive - EPITA/ING2/URSI/FNAC DARTY/Urbanisation shared/Urbanisation shared all/Matrice des flux.xlsx'
                              , data_only=True)
    # print(workSheet.get_sheet_names())
    matrix = workSheet["Matrice des flux"]
    cell_range = matrix['M1':'M43']
    # sheet_ranges = matrix['Graphviz']
    # print(sheet_ranges['D18'].value)
    for cell in cell_range:
        print(cell[0].value)

def render_graph():
    update = datetime.datetime.now()
    title = 'Matrice des flux (' + update.strftime('%Y/%m/%d %H:%M:%S') + ')'
    dot = Digraph(comment=title)
    dot.node('A', 'Alice')
    dot.node('B', 'Bob')
    dot.node('C', 'Cloe')
    dot.edges(['AB', 'AC'])
    dot.edge('B', 'C', constraint='false')
    print(dot.source)
    dot.render('matrix.gv', view=True)


if __name__ == "__main__":
    main()

