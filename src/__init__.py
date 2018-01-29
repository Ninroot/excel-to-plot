import datetime
import sys

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


def main(argv):
    if len(argv) == 0:
        generate()
    elif len(argv) == 1:
        generate(file_src=argv[0])
    else:
        generate(file_src=argv[0], dir_dest=argv[1])


def generate(file_src="Matrix.xlsx", dir_dest="."):
    workSheet = load_workbook(
        filename=file_src,
        data_only=True)
    # print(workSheet.get_sheet_names())
    matrix = workSheet["Matrice des flux"]

    update = datetime.datetime.now()
    title = 'Matrice des flux (màj : ' + update.strftime('%Y/%m/%d %H:%M:%S') + ')'
    dot = Digraph(comment='https://github.com/Ninroot/excel-to-plot', name=title, format='pdf')
    dot.body.append(get_legend())
    dot.graph_attr['fontsize'] = '20'
    dot.graph_attr['fontname'] = 'calibri'
    dot.node_attr['fontsize'] = '14'
    dot.node_attr['fontname'] = 'calibri'
    dot.edge_attr['fontsize'] = '14'
    dot.edge_attr['fontname'] = 'calibri'
    dot.body.append('\tlabelloc="t";\n\tlabel="' + title + '";')
    # dot.engine = 'circo'
    # dot.engine = 'sfdp'

    apps = ['Backoffice',
            'Caisse',
            'Fidélité',
            'Référentiel',
            'Monétique',
            'eCommerce',
            'Entrepôt',
            'BI',
            'Réappro',
            'Banque',
            'Fournisseur',
            'Client'
            ]

    for app in apps:
        dot.node(app, app, color=get_code_color_by_app(app), style='filled')

    # bound of the table, increase the max value if necessary
    for r in range(2, 70):
        row = matrix['A' + str(r) + ':' + 'N' + str(r)]
        flux = Flux(row)
        if not flux.is_valid():
            continue
        # print(flux)
        # dot.attr('node', shape='rarrow')
        # dot.edge_attr.update(arrowhead='vee', arrowsize='2')
        dot.edge(flux.src, flux.dst, flux.get_label(),
                 color=get_code_color_by_app(flux.src),
                 style=flux.get_style(),
                 arrowhead=flux.get_arrow_head(),
                 penwidth='2',
                 labelfontcolor='red'
                 )

    print(dot.source)
    dot.render(filename='matrix.gv',
               view=True,
               cleanup=True,
               directory=dir_dest)

    print("render in " + dir_dest + " directory")


class Flux:
    id = 0
    src = ""
    dst = ""
    title = ""
    route = ""
    progression = 0
    init = Initiator.PUSH
    temp = Temporality.SYNC

    def __init__(self, row):
        if row is None or row[0] is None:
            return None
        if row[0][0].value is not None:
            self.src = row[0][0].value
        if row[0][10].value is not None:
            self.route = row[0][10].value
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
            return 'dashed'
        return 'filled'

    def get_arrow_head(self):
        if self.init == Initiator.PULL:
            return "crow"
        return "vee"

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
        return '<<font color="{:s}"> [{:d}]{:s}({:.0f}%)<br/>{:s}</font>>'.format(get_code_color_by_app(self.src), self.id,
                                                                          self.title, self.progression, self.route)

    def is_valid(self):
        if self.src and self.dst:
            return True
        return False


def get_legend():
    legend = """
    rankdir=LR
    node [shape=plaintext]
    subgraph cluster_01 { 
      label = "Légende";
      key [label=<<table border="0" cellpadding="2" cellspacing="0" cellborder="0">
        <tr><td align="right" port="i0">Format</td></tr>
        <tr><td align="right" port="i1">Flux synchrone</td></tr>
        <tr><td align="right" port="i2">Flux asynchrone</td></tr>
        <tr><td align="right" port="i3">Push</td></tr>
        <tr><td align="right" port="i4">Pull</td></tr>
        </table>>]
      key2 [label=<<table border="0" cellpadding="2" cellspacing="0" cellborder="0">
        <tr><td align="right" port="i0">&nbsp;</td></tr>
        <tr><td port="i1">&nbsp;</td></tr>
        <tr><td port="i2">&nbsp;</td></tr>
        <tr><td port="i3">&nbsp;</td></tr>
        <tr><td port="i4">&nbsp;</td></tr>
        </table>>]
      key:i0:e -> key2:i0:w [color=white, label="[id du flux] titre du flux (pourcentage d'avancement)"]
      key:i1:e -> key2:i1:w [arrowhead=vee, style=filled, penwidth=2]
      key:i2:e -> key2:i2:w [arrowhead=vee, style=dashed, penwidth=2]
      key:i3:e -> key2:i3:w [arrowhead=vee, penwidth=2]
      key:i4:e -> key2:i4:w [arrowhead=crow, penwidth=2]
    }"""
    return legend


# http://www.color-hex.com/color-palette/200
def get_code_color_by_app(app):
    if app == "Backoffice":
        return "#e57f00"
    elif app == "Caisse":
        return "#645188"
    elif app == "eCommerce":
        return "#886451"
    elif app == "Réappro":
        return "#528881"
    elif app == "Fidélité":
        return "#5fc300"
    elif app == "BI":
        return "#c900a2"
    elif app == "Monétique":
        return "#0497df"
    elif app == "Entrepôt":
        return "#b8c300"
    elif app == "Référentiel":
        return "#0900ff"
    elif app == "Réappro":
        return "#4e17ff"
    elif app == "Banque":
        return "#e12637"
    elif app == "Fournisseur":
        return "#e12637"
    elif app == "Client":
        return "#e12637"
    return "#000000"


if __name__ == "__main__":
    main(sys.argv[1:])
