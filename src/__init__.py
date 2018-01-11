import datetime
from enum import auto

from openpyxl import load_workbook

from graphviz import Digraph

styles = {
    'async': {
        'style': 'dashed',
    },
    'sync': {
        'style': 'dashed',
    }
}

class Temporality:
    SYNC = "synchrone"
    ASYNC = "asynchrone"


class Initiator:
    PUSH = "push"
    PULL = "pull"


def main():
    workSheet = load_workbook(
        '/Users/ninroot/OneDrive - EPITA/ING2/URSI/FNAC DARTY/Urbanisation shared/Urbanisation shared all/Matrice des flux.xlsx'
        , data_only=True)
    # print(workSheet.get_sheet_names())
    matrix = workSheet["Matrice des flux"]

    update = datetime.datetime.now()
    title = 'Matrice des flux (màj : ' + update.strftime('%Y/%m/%d %H:%M:%S') + ')'
    dot = Digraph(comment='https://github.com/Ninroot/excel-to-plot', name=title, format='pdf')
    dot.graph_attr['fontsize'] = '20'
    dot.graph_attr['fontname'] = 'calibri'
    dot.body.append('\tlabelloc="t";\n\tlabel="' + title + '";')
    # dot.engine = 'circo'
    # dot.engine = 'sfdp'

    dot.edge_attr['fontsize'] = '14'
    dot.edge_attr['fontname'] = 'calibri'  # (fontsize='14')
    for r in range(2, 50):
        row = matrix['A' + str(r) + ':' + 'N' + str(r)]
        flux = Flux(row)
        if not flux.is_valid():
            continue
        print(flux)
        # dot.attr('node', shape='rarrow')
        # dot.edge_attr.update(arrowhead='vee', arrowsize='2')
        dot.edge(flux.src, flux.dst, flux.get_label(),
                 color=flux.get_code_color_by_app(),
                 style=flux.get_style(),
                 arrowhead=flux.get_arrow_head()
                 )


    # print(dot.source)
    dot.render(filename='matrix.gv',
               view=True,
               cleanup=True,
               # )
    directory='/Users/ninroot/OneDrive - EPITA/ING2/URSI/FNAC DARTY/Urbanisation shared/Urbanisation shared all/')


class Flux:
    id = 0
    src = ""
    dst = ""
    title = ""
    progression = 0
    init = Initiator.PUSH
    temp = Temporality.SYNC

    def __init__(self, row):
        if row is None or row[0] is None:
            return None
        if row[0][0].value is not None:
            self.src = row[0][0].value
        if row[0][1].value is not None:
            self.dst = row[0][1].value
        if row[0][2].value is not None:
            if row[0][2].value == "Push":
                self.init = Initiator.PUSH
            else:
                self.init = Initiator.PULL
        if row[0][3].value is not None:
            if row[0][3].value == "Synchrone":
                self.temp = Temporality.SYNC
            else:
                self.temp = Temporality.ASYNC
        if row[0][6].value is not None:
            self.title = row[0][6].value
        if row[0][11].value is not None:
            self.progression = row[0][11].value * 100
        if row[0][12].value is not None:
            self.id = row[0][12].value

    def get_style(self):
        if self.temp == Temporality.ASYNC:
            return 'filled'
        return 'dashed'

    def get_arrow_head(self):
        if self.init == Initiator.PULL:
            return "crow"
        return "vee"

    # http://www.color-hex.com/color-palette/200
    def get_code_color_by_app(self):
        if self.src == "Caisse":
            return "#885159"
        elif self.src == "Backoffice":
            return "#645188"
        elif self.src == "eCommerce":
            return "#886451"
        elif self.src == "Réappro":
            return "#528881"
        elif self.src == "Fidélité":
            return "#4d7358"
        elif self.src == "BI":
            return "#4d5d53"
        elif self.src == "Monétique":
            return "#00308f"
        elif self.src == "Entrepôt":
            return "#0900ff"
        # elif self.src == "":
        #     return ""
        # elif self.src == "":
        #     return ""

    def get_progression_code_color(self):
        if self.progression < 25:
            return '#d20000'
        if self.progression < 50:
            return '#d25f00'
        if self.progression < 75:
            return '#c6d200'
        return '#00d20e'

    def __str__(self):
        return "src: " + self.src + " dst: " + self.dst + " title: " + self.title + " temp: " + self.temp

    def get_label(self):
        return '[' + str(self.id) + ']' + self.title + '(' + str(self.progression) + '%)'

    def is_valid(self):
        if self.src and self.dst:
            return True
        return False


if __name__ == "__main__":
    main()
