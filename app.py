from flask import Flask, render_template
import zipfile
import xml.etree.ElementTree as ET

APP_XLSX = 'DISARM_MASTER_DATA/DISARM_FRAMEWORKS_MASTER.xlsx'
NAMESPACE = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

app = Flask(__name__)


def read_sheet(path, sheet_id):
    """Return sheet data as list of rows (lists of values)."""
    with zipfile.ZipFile(path) as z:
        # load shared strings
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('ns:si', NAMESPACE):
                text = ''.join(el.text or '' for el in si.findall('.//ns:t', NAMESPACE))
                shared.append(text)
        sheet_xml = f'xl/worksheets/sheet{sheet_id}.xml'
        sheet_root = ET.fromstring(z.read(sheet_xml))
        data = []
        for row in sheet_root.findall('.//ns:row', NAMESPACE):
            row_data = []
            for c in row.findall('ns:c', NAMESPACE):
                v = c.find('ns:v', NAMESPACE)
                if v is None:
                    value = ''
                else:
                    value = v.text
                    if c.attrib.get('t') == 's':
                        value = shared[int(value)]
                row_data.append(value)
            data.append(row_data)
        return data


def load_data():
    # According to workbook.xml, sheet6 is "techniques" and sheet10 is "countermeasures"
    red_raw = read_sheet(APP_XLSX, 6)
    blue_raw = read_sheet(APP_XLSX, 10)

    red_header = red_raw[0]
    blue_header = blue_raw[0]

    red = [dict(zip(red_header, row)) for row in red_raw[1:] if any(row)]
    blue = [dict(zip(blue_header, row)) for row in blue_raw[1:] if any(row)]
    return red, blue


@app.route('/')
def index():
    red, blue = load_data()
    return render_template('index.html', red=red, blue=blue)


if __name__ == '__main__':
    app.run(debug=True)
