VERSION = '0.1.1'
print(f'Slope to KML V{VERSION}\n')

import sys, os
import shutil
import json
from webbrowser import get
import pandas as pd
from shapely.geometry import Point, Polygon
from rich import print
import warnings
with warnings.catch_warnings():
    warnings.filterwarnings("ignore")
    from fastkml import kml, styles # ignore missing lxml

app_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))

def generateFilter():
    shutil.copy2(os.path.join(app_path, 'slope_filter.csv'), 'slope_filter.csv')

if __name__ == "__main__":
    k = kml.KML()
    d = kml.Document()

    print(f'Loading styles...')
    with open(os.path.join(app_path, 'styles.json'), 'r') as fil:
        styleDict = json.load(fil)
        for sk, sv in styleDict.items():
            d.append_style(styles.Style(styles=[styles.IconStyle(color=sv['iconColour'])], id=f'i{sk}'))
            d.append_style(styles.Style(styles=[styles.LineStyle(color=sv['shapeColour'])], id=f's{sk}'))

    print(f'Loading slope database...')
    with open(os.path.join(app_path, 'slopes.json'), 'r') as fil:
        slopes = json.load(fil)

    try:
        print(f'Loading filters from [bold]slope_filter.csv[/]...')
        filter = pd.read_csv('slope_filter.csv') # id, slopenum, description, category
        print(f'[green]loaded {len(filter)} filters[/]')
    except FileNotFoundError:
        print(f'[yellow][bold]slope_filter.csv[/] was not found, please edit and try again.[/]')
        generateFilter()
    except Exception:
        print(f'[red][bold]slope_filter.csv[/] contains errors, please edit and try again.[/]')
        generateFilter()
    else:
        for _, f in filter.iterrows():
            refnum, slopenum, description, batchnum = f.refnum, str(f.slopenum).replace(" ", ""), f.description, f.batchnum if f.batchnum else 0
            if slopenum not in slopes:
                print(f'[red]✕ {slopenum} not found![/]')
                continue
            try:
                poly = Polygon(slopes[slopenum])
                cntr = list(poly.centroid.coords)[0]

                pmpoly = kml.Placemark(name=f'{refnum} | {slopenum} | {description}', description=f'Batch {batchnum}', styleUrl=f"#s{batchnum}")
                pmpoly.geometry = kml.Geometry(geometry=poly)
                d.append(pmpoly)

                pmpoint = kml.Placemark(name=f'{refnum} | {slopenum} | {description}', description=f'Batch {batchnum}', styleUrl=f"#i{batchnum}")
                pmpoint.geometry = Point(cntr[0], cntr[1])
                d.append(pmpoint)

                print(f'[green]✓ {slopenum}: added to kml[/]')
            except Exception as e:
                print(f'[red]✕ {slopenum}: {e}[/]')

        k.append(d)
        with open('slope.kml', 'w', encoding='utf-8') as fil:
            fil.write(k.to_string().replace('kml:', ''))

        print(f'[green]kml written to [bold]slope.kml[/]![/]')
    finally:
        input(f'Press any key to exit...')