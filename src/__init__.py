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
    dot = Digraph(comment=title, name=title)
    # dot.engine = 'circo'
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

    for r in range(2, 50):
        row = matrix['A' + str(r) + ':' + 'N' + str(r)]
        flux = Flux(row)
        if not flux.is_valid():
            continue
        print(flux)
        dot.edge(flux.src, flux.dst, flux.get_label())

    # print(dot.source)
    dot.render(filename='matrix.gv',
               view=True,
               directory='/Users/ninroot/OneDrive - EPITA/ING2/URSI/FNAC DARTY/Urbanisation shared/Urbanisation shared all/')


class Flux:
    id = 0
    src = ""
    dst = ""
    title = ""
    progression = 0

    def __init__(self, row):
        if row is None or row[0] is None:
            return None
        if row[0][0].value is not None:
            self.src = row[0][0].value
        if row[0][1].value is not None:
            self.dst = row[0][1].value
        if row[0][6].value is not None:
            self.title = row[0][6].value
        if row[0][11].value is not None:
            self.progression = row[0][11].value
        if row[0][12].value is not None:
            self.id = row[0][12].value

    def __str__(self):
        return "src: " + self.src + " dst: " + self.dst + " title: " + self.title

    def get_label(self):
        return '[' + str(self.id) + ']' + self.title + '(' + str(self.progression) + '%)'

    def is_valid(self):
        if self.src and self.dst:
            return True
        return False


if __name__ == "__main__":
    main()
