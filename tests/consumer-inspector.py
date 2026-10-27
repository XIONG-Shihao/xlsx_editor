"""Negative controls for the finite consumer oracle; all packages are temporary."""
import copy
import importlib.util
import pathlib
import tempfile
import unittest
import zipfile

spec = importlib.util.spec_from_file_location('inspector', pathlib.Path(__file__).parents[1] / 'scripts/test/xlsx-consumer-inspect.py')
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)
stream_spec = importlib.util.spec_from_file_location('stream_oracle', pathlib.Path(__file__).parents[1] / 'scripts/test/xlsx-stream-preservation.py')
stream_oracle = importlib.util.module_from_spec(stream_spec)
stream_spec.loader.exec_module(stream_oracle)
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def package():
    return {
        'xl/workbook.xml': f'<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="sheet"/></sheets><pivotCaches><pivotCache cacheId="1" r:id="cache"/></pivotCaches><externalReferences><externalReference r:id="external"/></externalReferences></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="sheet" Target="worksheets/sheet1.xml"/><Relationship Id="cache" Target="pivotCache/cache.xml"/><Relationship Id="external" Target="externalLinks/link.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>9007199254740993</v></c><c r="B1"><v>-0</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><f>NA()</f><v>#N/A</v></c><c r="E1" t="inlineStr"><is><r><t>東京</t></r><rPh sb="0" eb="2"><t>とうきょう</t></rPh></is></c></row></sheetData></worksheet>',
        'xl/charts/chart.xml': f'<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:lineChart><c:ser><c:val><c:numRef><c:f>Sheet1!A1:A2</c:f><c:numCache><c:pt idx="0"><c:v>9007199254740993</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>',
        'xl/pivotTables/table.xml': f'<pivotTableDefinition xmlns="{S}" name="Pivot1" cacheId="1"><location ref="A4:B6"/><dataFields><dataField fld="1" subtotal="sum"/></dataFields></pivotTableDefinition>',
        'xl/pivotCache/cache.xml': f'<pivotCacheDefinition xmlns="{S}"><cacheSource type="worksheet"><worksheetSource sheet="Sheet1" ref="A1:B3"/></cacheSource><cacheFields><cacheField name="Label"/><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>',
        'xl/externalLinks/_rels/link.xml.rels': '<Relationships><Relationship Target="file:///source.xlsx" TargetMode="External"/></Relationships>',
        'xl/media/image.png': b'png', 'xl/embeddings/object.bin': b'opaque',
        'xl/drawings/shape.xml': f'<xdr:wsDr xmlns:xdr="{D}" xmlns:a="{A}"><xdr:oneCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:row>3</xdr:row></xdr:from><xdr:ext cx="100" cy="200"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr name="sample" descr="Alt"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr><xdr:txBody><a:p><a:r><a:rPr sz="1400"><a:latin typeface="Arial"/></a:rPr><a:t>Keep text</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:oneCellAnchor></xdr:wsDr>',
        'xl/comments/comment.xml': f'<comments xmlns="{S}"><authors><author>Alice</author></authors><commentList><comment ref="A1" authorId="0"><text><t>Keep note</t></text></comment></commentList></comments>',
        'xl/worksheets/_rels/sheet1.xml.rels': f'<Relationships><Relationship Type="{R}/comments" Target="../comments/comment.xml"/></Relationships>',
    }


class ConsumerInspector(unittest.TestCase):
    def compare(self, mutate=lambda p: None):
        original = package()
        returned = copy.deepcopy(original)
        mutate(returned)
        with tempfile.TemporaryDirectory() as directory:
            paths = []
            for name, parts in [('original', original), ('returned', returned)]:
                file = pathlib.Path(directory) / (name + '.xlsx')
                with zipfile.ZipFile(file, 'w') as archive:
                    for part, content in parts.items():
                        archive.writestr(part, content)
                paths.append(file)
            return oracle.compare(*(oracle.inspect(file) for file in paths))

    def replace(self, part, before, after):
        return self.compare(lambda p: p.update({part: p[part].replace(before, after)}))

    def test_identical_passes(self):
        self.assertTrue(self.compare()['passed'])

    def test_numeric_precision_and_sign(self):
        part = 'xl/worksheets/sheet1.xml'
        for before, after in [('9007199254740993', '9007199254740992'), ('<v>-0</v>', '<v>0</v>'), ('t="b"', 't="n"')]:
            with self.subTest(before=before):
                self.assertFalse(self.replace(part, before, after)['passed'])
        self.assertTrue(self.replace(part, '<v>-0</v>', '<v>-0.000</v>')['passed'])

    def test_chart_deletion_references_and_precision(self):
        part = 'xl/charts/chart.xml'
        self.assertFalse(self.compare(lambda p: p.pop(part))['passed'])
        for before, after in [('A1:A2', 'B1:B2'), ('9007199254740993', '9007199254740992')]:
            with self.subTest(before=before):
                self.assertFalse(self.replace(part, before, after)['passed'])

    def test_pivot_deletion_source_and_aggregate(self):
        self.assertFalse(self.compare(lambda p: p.pop('xl/pivotTables/table.xml'))['passed'])
        self.assertFalse(self.replace('xl/pivotTables/table.xml', 'subtotal="sum"', 'subtotal="count"')['passed'])
        self.assertFalse(self.replace('xl/pivotCache/cache.xml', 'A1:B3', 'A1:B4')['passed'])

    def test_unused_external_target_loss(self):
        self.assertFalse(self.replace('xl/workbook.xml', '<externalReferences><externalReference r:id="external"/></externalReferences>', '')['passed'])
        self.assertFalse(self.replace('xl/externalLinks/_rels/link.xml.rels', 'source.xlsx', 'other.xlsx')['passed'])

    def test_cached_type_change_is_reported(self):
        report = self.replace('xl/worksheets/sheet1.xml', 'r="D1" t="e"', 'r="D1" t="str"')
        self.assertEqual(report['formulaCacheChanges'], 1)
        self.assertTrue(report['passed'])  # Recalculated caches are explicitly outside the literal pass gate.

    def test_phonetic_annotations_are_not_visible_value(self):
        self.assertTrue(self.replace('xl/worksheets/sheet1.xml', '<rPh sb="0" eb="2"><t>とうきょう</t></rPh>', '')['passed'])

    def test_external_renumbering_preserves_literal_and_selector_text(self):
        external = {'1': ['file:///a.xlsx'], '2': ['file:///a.xlsx']}
        self.assertEqual(oracle.normalized_formula('[1]Sheet!A1', external), oracle.normalized_formula('[2]Sheet!A1', external))
        for formula in ['"[1]Sheet!A1"', 'Table[1]', 'SUM(Table[1])+Sheet!A1']:
            self.assertEqual(oracle.normalized_formula(formula, external), formula)
        self.assertEqual(oracle.decode('_x005F_x0041_'), '_x0041_')
        self.assertEqual(oracle.decode('_xD83D__xDE00_'), '😀')

    def test_missing_and_unexpected_literals(self):
        part = 'xl/worksheets/sheet1.xml'
        self.assertFalse(self.replace(part, '<c r="B1"><v>-0</v></c>', '')['passed'])
        self.assertFalse(self.replace(part, '</row>', '<c r="F1"><v>7</v></c></row>')['passed'])

    def test_opaque_payload_loss(self):
        for part in ['xl/media/image.png', 'xl/embeddings/object.bin']:
            with self.subTest(part=part):
                self.assertFalse(self.compare(lambda p: p.pop(part))['passed'])

    def test_shape_deletion_text_geometry_and_font(self):
        part = 'xl/drawings/shape.xml'
        self.assertFalse(self.compare(lambda p: p.pop(part))['passed'])
        for before, after in [('Keep text', 'Lost text'), ('prst="rect"', 'prst="ellipse"'), ('typeface="Arial"', 'typeface="Times"'), ('descr="Alt"', 'descr=""')]:
            with self.subTest(before=before):
                self.assertFalse(self.replace(part, before, after)['passed'])

    def test_note_text_author_and_worksheet_binding(self):
        for part, before, after in [('xl/comments/comment.xml', 'Keep note', 'Lost note'), ('xl/comments/comment.xml', 'Alice', 'Bob'), ('xl/worksheets/_rels/sheet1.xml.rels', '/comments"', '/drawing"')]:
            with self.subTest(part=part):
                self.assertFalse(self.replace(part, before, after)['passed'])

    def test_explicit_theme_loss_and_font_change_are_detected(self):
        p = package()
        p['xl/_rels/workbook.xml.rels'] = p['xl/_rels/workbook.xml.rels'].replace('</Relationships>', f'<Relationship Id="theme" Type="{R}/theme" Target="theme/theme.xml"/></Relationships>')
        p['xl/theme/theme.xml'] = f'<a:theme xmlns:a="{A}"><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="112233"/></a:dk1></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Georgia"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'
        with tempfile.TemporaryDirectory() as directory:
            def inspect(parts):
                file = pathlib.Path(directory)/'theme.xlsx'
                with zipfile.ZipFile(file, 'w') as z:
                    for name, data in parts.items():z.writestr(name, data)
                return oracle.inspect(file)
            original = inspect(p)
            p['xl/theme/theme.xml'] = p['xl/theme/theme.xml'].replace('Georgia', 'Times')
            self.assertFalse(oracle.compare(original, inspect(p))['passed'])
            p['xl/_rels/workbook.xml.rels'] = p['xl/_rels/workbook.xml.rels'].replace(f'<Relationship Id="theme" Type="{R}/theme" Target="theme/theme.xml"/>', '')
            self.assertFalse(oracle.compare(original, inspect(p))['passed'])

    def test_shadow_defaults_preserve_color_opacity_and_nondefault_transform(self):
        node = oracle.E.fromstring(f'<a:outerShdw xmlns:a="{A}" sx="100000" sy="100000" kx="0" ky="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>')
        minimal = copy.deepcopy(node)
        minimal.attrib.clear()
        self.assertEqual(oracle.shadow_key(node), oracle.shadow_key(minimal))
        minimal.find('{'+A+'}srgbClr/{'+A+'}alpha').set('val', '90000')
        self.assertNotEqual(oracle.shadow_key(node), oracle.shadow_key(minimal))
        minimal.attrib['sx'] = '110000'
        self.assertNotEqual(oracle.shadow_key(node), oracle.shadow_key(minimal))

    def test_header_text_and_named_defaults(self):
        self.assertEqual(oracle.header_sections('&LPlain && text'), oracle.header_sections('&L&"Calibri"&11&K112233Plain && text'))
        self.assertNotEqual(oracle.header_sections('&LPlain'), oracle.header_sections('&LChanged'))
        self.assertEqual(oracle.header_sections('&R&P'), {'unqualified': '&R&P'})
        self.assertEqual(oracle.name_attributes(oracle.E.fromstring('<definedName name="N" hidden="false" function="0"/>')), {'name': 'N'})
        self.assertNotEqual(oracle.name_attributes(oracle.E.fromstring('<definedName name="N" hidden="true"/>')), {'name': 'N'})
        self.assertEqual(oracle.normalized_formula("'Sheet1'!A1", {}), 'Sheet1!A1')
        self.assertEqual(oracle.normalized_formula('"\'Sheet1\'!A1"', {}), '"\'Sheet1\'!A1"')


class StreamingPointInspector(unittest.TestCase):
    def compare(self, mutate=lambda p: None, allow_caches=False, prepare=lambda p: None):
        original = package()
        prepare(original)
        returned = copy.deepcopy(original)
        mutate(returned)
        with tempfile.TemporaryDirectory() as directory:
            paths = []
            for name, parts in [('original', original), ('returned', returned)]:
                file = pathlib.Path(directory) / (name + '.xlsx')
                with zipfile.ZipFile(file, 'w') as archive:
                    for part, content in parts.items():
                        archive.writestr(part, content)
                paths.append(file)
            return stream_oracle.compare(*(str(p) for p in paths), 'xl/worksheets/sheet1.xml', ['A1'], allow_caches)

    def test_target_change_is_allowed_but_untouched_cells_and_envelopes_are_not(self):
        part = 'xl/worksheets/sheet1.xml'
        self.assertTrue(self.compare(lambda p: p.update({part: p[part].replace('<v>9007199254740993</v>', '<v>7</v>')}))['passed'])
        for before, after in [('r="A1"', 'r="A1" cm="1"'), ('<v>-0</v>', '<v>8</v>'), ('<row r="1">', '<row r="1" hidden="1">')]:
            with self.subTest(before=before):
                self.assertFalse(self.compare(lambda p: p.update({part: p[part].replace(before, after)}))['passed'])

    def test_cache_permission_never_changes_authored_formula_or_active_attributes(self):
        part = 'xl/worksheets/sheet1.xml'
        cache = lambda p: p.update({part: p[part].replace('<v>#N/A</v>', '<v>#VALUE!</v>')})
        self.assertFalse(self.compare(cache)['passed'])
        self.assertTrue(self.compare(cache, True)['passed'])
        self.assertFalse(self.compare(lambda p: p.update({part: p[part].replace('NA()', 'SUM(A1)')}), True)['passed'])
        prepare = lambda p: p.update({part: p[part].replace('<f>NA()', '<f aca="true">NA()')})
        self.assertFalse(self.compare(lambda p: p.update({part: p[part].replace(' aca="true"', '')}), True, prepare)['passed'])
        prepare = lambda p: p.update({part: p[part].replace('<f>NA()', '<f aca="false">NA()')})
        self.assertTrue(self.compare(lambda p: p.update({part: p[part].replace(' aca="false"', '')}), True, prepare)['passed'])

    def test_package_graph_and_unknown_calculation_metadata_stay_protected(self):
        self.assertFalse(self.compare(lambda p: p.pop('xl/embeddings/object.bin'))['passed'])
        part = 'xl/workbook.xml'
        prepare = lambda p: p.update({part: p[part].replace('</workbook>', '<calcPr vendorSetting="keep"/></workbook>')})
        self.assertFalse(self.compare(lambda p: p.update({part: p[part].replace(' vendorSetting="keep"', '')}), True, prepare)['passed'])

    def literal_comparison(self, mutate=lambda p: None, keep_formula=False):
        original = package()
        part = 'xl/worksheets/sheet1.xml'
        if not keep_formula:
            original[part] = original[part].replace('<f>NA()</f>', '')
        returned = copy.deepcopy(original)
        mutate(returned)
        with tempfile.TemporaryDirectory() as directory:
            paths = []
            for name, parts in [('original', original), ('returned', returned)]:
                file = pathlib.Path(directory) / (name + '.xlsx')
                with zipfile.ZipFile(file, 'w') as archive:
                    for member, content in parts.items():
                        archive.writestr(member, content)
                paths.append(file)
            return stream_oracle.compare_literal_data(*(str(p) for p in paths))

    def test_large_data_profile_compares_decoded_strings_and_exact_numbers(self):
        self.assertTrue(self.literal_comparison()['passed'])
        part = 'xl/worksheets/sheet1.xml'
        def shared(p):
            p[part] = p[part].replace('<c r="E1" t="inlineStr"><is><r><t>東京</t></r><rPh sb="0" eb="2"><t>とうきょう</t></rPh></is></c>', '<c r="E1" t="s"><v>0</v></c>')
            p['xl/sharedStrings.xml'] = f'<sst xmlns="{S}"><si><t>東京</t></si></sst>'
        self.assertTrue(self.literal_comparison(shared)['passed'])
        lost = self.literal_comparison(lambda p: p.update({part: p[part].replace('9007199254740993', '9007199254740992')}))
        self.assertFalse(lost['passed'])
        self.assertTrue(lost['binary64DataRetained'])

    def test_large_literal_profile_refuses_to_invent_shared_formula_equivalence(self):
        with self.assertRaises(ValueError):
            self.literal_comparison(keep_formula=True)



# This class is discovered when the module is imported by unittest.
class CellLifecycleOracle(unittest.TestCase):
    def setUp(self):
        spec=importlib.util.spec_from_file_location('package_oracle',pathlib.Path(__file__).parents[1]/'scripts/test/xlsx-package-check.py')
        self.checker=importlib.util.module_from_spec(spec);spec.loader.exec_module(self.checker)
        self.old=f'<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>9</v></c></row></sheetData></worksheet>'
        self.new=f'<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="C1" t="inlineStr"><is><t>new</t></is></c></row></sheetData></worksheet>'
    def failures(self,old=None,new=None,added=('C1',),cleared=('B1',)):
        return self.checker.strict_cell_envelopes(old or self.old,new or self.new,['B1','C1'],{}, {},(),added,cleared)
    def test_declared_add_and_clear_pass(self):self.assertEqual(self.failures(),[])
    def test_missing_add_or_clear_authority_fails(self):
        self.assertTrue(self.failures(added=()));self.assertTrue(self.failures(cleared=()))
    def test_new_cell_annotations_and_formulas_fail(self):
        self.assertTrue(self.failures(new=self.new.replace('r="C1"','r="C1" xmlns:u="urn:foreign" u:marker="bad"')))
        self.assertTrue(self.failures(new=self.new.replace('<is><t>new</t></is>','<f>1+1</f><v>2</v>')))
    def test_clearing_annotated_source_or_unrelated_row_metadata_fails(self):
        self.assertTrue(self.failures(old=self.old.replace('r="B1"','r="B1" xmlns:u="urn:foreign" u:marker="keep"')))
        self.assertTrue(self.failures(new=self.new.replace('r="1"','r="1" hidden="1"',1)))
    def test_added_row_cannot_hide_foreign_metadata(self):
        new=self.new.replace('</sheetData>','<row r="2" hidden="1"/></sheetData>')
        self.assertTrue(self.failures(new=new))

if __name__ == '__main__':
    unittest.main()
