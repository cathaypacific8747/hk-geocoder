VERSION = '0.2'

import sys, os
import shutil
import json
import pandas as pd
from styleframe import StyleFrame
import numpy as np
from rich.console import Console
import kml

console = Console()

console.print(f'[light_cyan1 bold]Slope to KML v{VERSION}[/]', highlight=False)
console.print(f'[grey50]Source Code: https://www.github.com/cathaypacific8747/hk-slope-kml[/]')
console.print(f'[grey50]License: MIT[/]\n')

app_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))

if __name__ == "__main__":
    if not os.path.isfile('filter.xlsx'):
        shutil.copy2(os.path.join(app_path, 'filter.xlsx'), 'filter.xlsx')
        console.print(f'        INFO: A new filter file is created in the current directory.')
        console.print(f'              Please edit it and run this program again.')
        sys.exit(0)

    kmlFile = kml.KML()
    def transform(cell):
        col = cell.style.column
        argb = str(cell.style.fill.fgColor.rgb)
        if col == 8: # batchNum, return nan if nan
            return np.nan if pd.isna(cell.value) else str(cell.value)
        elif col == 9: # iconColour
            return f'ff{argb[6:8]}{argb[4:6]}{argb[2:4]}'
        elif col == 10: # shapeColour
            return f'aa{argb[6:8]}{argb[4:6]}{argb[2:4]}'
        return cell
    
    console.print(f'        INFO: Loading slope database...')
    with open(os.path.join(app_path, 'slopes.json'), 'r') as fil:
        slopes = json.load(fil)

    console.print(f'        INFO: Reading filter.xlsx...')
    df = StyleFrame.read_excel(
        path='filter.xlsx',
        read_style=True,
        use_openpyxl_styles=True,
    ).applymap(transform)
    
    console.print(f'        INFO: Adding icon and shape styles...')
    s_df = df.iloc[2:, [7,8,9]].dropna()
    s_df.columns = ['batchNum', 'iconColour', 'shapeColour']
    validBatchNums = list(s_df.batchNum)
    for i, row in s_df.iterrows():
        batchNum, iconColour, shapeColour = row
        kmlFile.addStyle('icon', f'i{batchNum}', iconColour)
        kmlFile.addStyle('shape', f's{batchNum}', shapeColour)

    console.print(f'        INFO: Adding entries...')
    m_df = df.iloc[2:, 1:6].dropna()
    m_df.columns = ['refNum', 'value', 'vtype', 'description', 'batchNum']

    for _, f in m_df.iterrows():
        refnum = str(f.refNum)
        value = str(f.value).replace(' ', '')
        vtype = str(f.vtype)
        description = str(f.description)
        batchNum = str(f.batchNum) if str(f.batchNum) in validBatchNums else '0'

        if vtype == 'slopeno':
            if value not in slopes:
                console.print(f'[red]{value:>12}: Not found.[/]')
                continue
            try:
                kmlFile.addEntry(
                    polyCoords=slopes[value],
                    header=f'{refnum} | {value} | {description}',
                    batchNum=batchNum,
                )
                console.print(f'[chartreuse3]{value:>12}: Done[/]')
            except Exception as e:
                console.print(f'[red]{value:>12}: Critical Error - {e}[/]')
        else:
            console.print(f'[yellow]{value:>12}: Search by lotno will be implemented in future versions.[/]')

    kmlFile.save('slope.kml')
    console.print(f'[chartreuse3]KML successfully generated to [bold]slope.kml[/].[/]')

    input(f'Press any key to exit...')
    sys.exit(0)