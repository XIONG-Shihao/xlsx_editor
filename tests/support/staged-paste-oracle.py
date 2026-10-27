"""Independent streaming check for the dense numeric-one paste fixture."""
import hashlib
import json
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def column(number):
    result = ''
    number += 1
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def file_hash(filename):
    digest = hashlib.sha256()
    with open(filename, 'rb') as file:
        for chunk in iter(lambda: file.read(65536), b''):
            digest.update(chunk)
    return digest.hexdigest()


def inspect(source, saved, rows, columns):
    worksheet = 'xl/worksheets/sheet1.xml'
    count = 0
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(saved) as exported:
        if len(exported.namelist()) != len(set(exported.namelist())):
            raise ValueError('Duplicate saved package member')
        if set(original.namelist()) != set(exported.namelist()):
            raise ValueError('Package member set changed')
        for name in original.namelist():
            if name != worksheet and original.read(name) != exported.read(name):
                raise ValueError('Unchanged package member differs: ' + name)
        with exported.open(worksheet) as file:
            for event, element in ET.iterparse(file, events=('end',)):
                if element.tag == NS + 'c':
                    expected = column(count % columns) + str(count // columns + 1)
                    if element.get('r') != expected:
                        raise ValueError('Missing, duplicated or unordered saved cell: ' + expected)
                    if element.get('t') not in (None, 'n'):
                        raise ValueError('Saved numeric type changed at ' + expected)
                    value = element.find(NS + 'v')
                    if value is None or element.find(NS + 'f') is not None:
                        raise ValueError('Saved numeric scalar changed at ' + expected)
                    if struct.pack('>d', float(value.text)) != struct.pack('>d', 1.0):
                        raise ValueError('Saved binary64 value changed at ' + expected)
                    count += 1
                if element.tag == NS + 'row':
                    element.clear()
    if count != rows * columns:
        raise ValueError('Saved cell count disagrees with the complete rectangle')
    return {'savedCellCount': count, 'savedNumericOneCount': count,
            'numericTypeVerified': True, 'binary64Verified': True,
            'allUnchangedPackageMembersPreserved': True,
            'sourceSha256': file_hash(source), 'exportSha256': file_hash(saved)}


if __name__ == '__main__':
    print(json.dumps(inspect(sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4]))))
