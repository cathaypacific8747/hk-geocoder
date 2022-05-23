import warnings
from shapely.geometry import Point, Polygon
with warnings.catch_warnings():
    warnings.filterwarnings("ignore")
    from fastkml import kml, styles # ignore missing lxml

class KML:
    def __init__(self):
        self.kml = kml.KML()
        self.document = kml.Document()
    
    def addStyle(self, styleType, styleTag, styleColour):
        ss = [styles.IconStyle(color=styleColour)] if styleType == 'icon' else [styles.LineStyle(color=styleColour), styles.PolyStyle(color=styleColour)]
        self.document.append_style(styles.Style(styles=ss, id=styleTag))
    
    def addEntry(self, polyCoords, header, batchNum=0):
        poly = Polygon(polyCoords)
        cntrLon, cntrLat = list(poly.centroid.coords)[0]

        pmpoint = kml.Placemark(name=header, description=f'Batch {batchNum}', styleUrl=f"#s{batchNum}") # shape
        pmpoint.geometry = kml.Geometry(geometry=poly)
        self.document.append(pmpoint)

        pmpoint = kml.Placemark(name=header, description=f'Batch {batchNum}', styleUrl=f"#i{batchNum}") # icon
        pmpoint.geometry = Point(cntrLon, cntrLat)
        self.document.append(pmpoint)

    def save(self, filename):
        self.kml.append(self.document)
        with open(filename, 'w', encoding='utf-8') as fil:
            fil.write(self.kml.to_string().replace('kml:', ''))