# hk-geocoder
Batch KML geocoding tool to ease local slope inspections and deliveries in Hong Kong.

Given the:
- [slope numbers](https://hkss.cedd.gov.hk/hkss/en/home/index.html), or
- [lot codes](https://www.landreg.gov.hk/en/public/si_web.htm)

in an Excel file, this program will attempt to geocode the location and aggregate matched polygons/centroids to a KML file, which cna then be uploaded to [Google My Maps](https://www.google.com/maps/d/) for easy and rapid navigation on mobile devices. Features can also be grouped by batches and coloured based on custom configurations.

Data source: [SIS](https://hkss.cedd.gov.hk/hkss/en/facts-and-figures/slope-information-system/sis/index.html), [HKMS2.0](https://www.hkmapservice.gov.hk/OneStopSystem/map-search), others.

## For normal usage:
1. Go to the [releases](https://github.com/cathaypacific8747/hk-geocoder/releases) section and download the precompiled exe.
2. Run `generator.exe` once to generate the required filter.
3. Add/edit entries and/or styles in `filter.xlsx`.
4. Re-run `generator.exe` to generate the KML file.


## For development or for Mac/Linux users:
```
$ pip3 install pipenv
$ pipenv install
$ pipenv run python3 generator.py
$ pipenv run pyinstaller generator.spec
```
Executable will be generated under `dist/hk-geocoder-vX.X.exe`