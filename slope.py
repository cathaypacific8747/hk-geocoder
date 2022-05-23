VERSION = '0.2'

from rich.console import Console
import sys, os
import shutil
import json
import pandas as pd
from styleframe import StyleFrame
import numpy as np
import kml
import requests
import base64
import urllib
import warnings

warnings.filterwarnings("ignore", category=DeprecationWarning) 
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
        value = str(f.value)
        vtype = str(f.vtype)
        description = str(f.description)
        batchNum = str(f.batchNum) if str(f.batchNum) in validBatchNums else '0'

        if vtype == 'slopeno':
            value = value.replace(' ', '')
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
        else: # lotno
            try:
                key = base64.b64decode(b'ZGQ5NzA3OTk5MTlmNDlmMzkyOWVhNmIyYjVkNDdjZjU=').decode('ascii')
                reqHeaders = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.67 Safari/537.36",
                    "Accept": "application/json",
                    "Accept-Language": "en-US,en;q=0.5",
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                    "Sec-Fetch-Dest": "empty",
                    "Sec-Fetch-Mode": "cors",
                    "Sec-Fetch-Site": "same-site",
                    "Pragma": "no-cache",
                    "Cache-Control": "no-cache",
                    "referer": "https://www.hkmapservice.gov.hk/",
                }
                r = requests.post(
                    f"https://api.hkmapservice.gov.hk/oss/services/MapAPI/Lot/GeocodeServer/findAddressCandidates?key={key}",
                    headers=reqHeaders,
                    data=f"SingleLine={urllib.parse.quote(value)}&key={key}&f=json&outFields=*"
                ).json()

                if not len(r['candidates']):
                    console.print(f'[red]{value:>12}: Not found.[/]')
                    continue
                else:
                    confidence = r['candidates'][0]['score']
                    r1 = requests.get(
                        f"https://api.hkmapservice.gov.hk/oss/services/OneStop/LPD/MapServer/1/query?key={key}&f=json&where=lotid%20=%20{r['candidates'][0]['attributes']['Ref_ID']}&returnGeometry=true&outSR=4326",
                        headers=reqHeaders
                    ).json()
                    kmlFile.addEntry(
                        polyCoords=r1['features'][0]['geometry']['rings'][0],
                        header=f'{refnum} | {value} | {description}',
                        batchNum=batchNum,
                    )
                    console.print(f'[chartreuse3]{value:>12}: Done with confidence {confidence}%[/]')
            except requests.exceptions.ConnectionError:
                console.print(f'[red]{value:>12}: Connection Error[/]')
                continue
            except Exception as e:
                console.print(f'[red]{value:>12}: Critical Error - {e}[/]')
                continue

    kmlFile.save('slope.kml')
    console.print(f'[chartreuse3]KML successfully generated to [bold]slope.kml[/].[/]')

    input(f'Press any key to exit...')
    sys.exit(0)