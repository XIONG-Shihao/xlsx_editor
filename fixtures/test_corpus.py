"""Checks the independent corpus and reference preparation, not a workbook adapter."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


def load(name, file):
    spec = importlib.util.spec_from_file_location(name, file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = load('xlsx_fixture_generator', Path(__file__).with_name('build-corpus.py'))
reference = load('xlsx_independent_reference', ROOT / 'scripts/test/xlsx-reference.py')


class CorpusTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temporary = tempfile.TemporaryDirectory(prefix='xlsx-corpus-check-')
        cls.directory = Path(cls.temporary.name)
        cls.manifest = generator.build(cls.directory / 'first')
        cls.keys = json.loads(Path(__file__).with_name('expectations.json').read_text())['fixtures']
        cls.inventory = {
            item['id']: reference.inspect(cls.directory / 'first' / item['file'], True)
            for item in cls.manifest
        }

    @classmethod
    def tearDownClass(cls):
        cls.temporary.cleanup()

    def test_generation_is_deterministic_and_keys_are_complete(self):
        self.assertEqual(self.manifest, generator.build(self.directory / 'second'))
        self.assertEqual(set(self.keys), set(self.inventory))

    def test_raw_values_formulas_and_intentional_caches_match_literal_keys(self):
        for name, key in self.keys.items():
            with self.subTest(fixture=name):
                failures = [item for item in reference.check_literals(self.inventory[name], key) if not item['passed']]
                self.assertEqual(failures, [])

    def test_scale_counts_and_dependency_depth_are_exact(self):
        for count in (10000, 100000):
            with self.subTest(count=count):
                item = self.inventory[f'scale-{count}']
                self.assertEqual(item['usedCells'], count)
                self.assertEqual(item['formulas'], count // 10)
        chain = self.inventory['dependency-chain']
        self.assertEqual(chain['usedCells'], 10000)
        self.assertEqual(chain['formulas'], 9999)
        self.assertEqual(chain['cells']['Chain!A10000']['formula'], 'A9999+1')

    def test_date_systems_and_excel_fictitious_day_are_kept_as_raw_serials(self):
        self.assertEqual(self.inventory['dates-1900']['dateSystem'], '1900')
        self.assertEqual(self.inventory['dates-1900']['cells']['Dates!A2']['rawValue'], '60')
        self.assertEqual(self.inventory['dates-1904']['dateSystem'], '1904')
        self.assertEqual(self.inventory['dates-1904']['cells']['Dates!A1']['rawValue'], '0')
        self.assertEqual(self.inventory['dates-1904']['cells']['Dates!A1']['style'], '1')

    def test_preservation_and_refusal_fixtures_really_have_the_declared_features(self):
        rich = self.inventory['preservation-rich']
        for feature in ('charts', 'imagesOrMedia', 'customXml'):
            self.assertGreater(rich['features'][feature], 0)
        self.assertEqual(rich['sheets'][1]['state'], 'hidden')
        self.assertEqual(rich['sheets'][0]['merges'], 1)
        self.assertEqual(rich['sheets'][0]['validations'], 1)
        self.assertEqual(self.inventory['pricing-schedule']['features']['tables'], 1)
        self.assertEqual(self.inventory['pricing-schedule']['namedRanges'], 1)
        self.assertEqual(self.inventory['budget']['sheets'][1]['state'], 'veryHidden')
        self.assertTrue(self.inventory['protected']['workbookProtected'])
        self.assertTrue(self.inventory['protected']['sheets'][0]['protected'])
        self.assertEqual(self.inventory['opaque-media']['members']['xl/media/opaque.bin']['bytes'], 16 * 1024 * 1024)
        for name, key in self.keys.items():
            for part in key.get('preservation', {}).get('requiredByteIdentical', []):
                with self.subTest(fixture=name, part=part):
                    self.assertIn(part, self.inventory[name]['members'])
                    self.assertNotIn(part, key['preservation']['allowedChanged'])

    def test_reference_preparation_forces_calculation_without_editing_originals(self):
        source = self.directory / 'first/calculation-states.xlsx'
        original_hash = reference.sha(source)
        target = self.directory / 'recalculation.xlsx'
        reference.write_scenario(source, target, [])
        prepared = reference.inspect(target, True)
        self.assertEqual(reference.sha(source), original_hash)
        for cell in prepared['cells'].values():
            if cell['formula']:
                self.assertIsNone(cell['rawValue'])
        self.assertEqual(prepared['cells']['Diagnostics!E1']['rawValue'], '9007199254740993')
        self.assertEqual(prepared['cells']['Diagnostics!E3']['value'], '=SUM(A1:A2)')

    def test_reference_scenario_applies_handwritten_inputs_and_keeps_opaque_parts(self):
        source = self.directory / 'first/pricing-schedule.xlsx'
        original_hash = reference.sha(source)
        target = self.directory / 'scenario.xlsx'
        key = self.keys['pricing-schedule']
        reference.write_scenario(source, target, key['scenario']['edits'])
        prepared = reference.inspect(target, True)
        self.assertEqual(reference.sha(source), original_hash)
        self.assertEqual(prepared['cells']['Pricing!B2']['value'], 3)
        self.assertEqual(prepared['cells']['Pricing!C2']['value'], 150)
        self.assertEqual(prepared['cells']['Pricing!D5']['formula'], 'SUM(PricingItems[Subtotal])')
        original = self.inventory['pricing-schedule']
        for part in ('xl/tables/table1.xml', 'xl/styles.xml'):
            self.assertEqual(prepared['members'][part], original['members'][part])

    def test_reference_engine_cannot_be_an_exact_precision_oracle(self):
        key = {'cells': {'Precision!A1': {'rawValue': '9007199254740993'}}}
        rounded = {'cells': {'Precision!A1': {'value': 9007199254740990, 'rawValue': '9007199254740990', 'formula': None}}}
        self.assertFalse(reference.check_literals(rounded, key)[0]['passed'])
        observed = reference.check_literals(rounded, key, calculated=True)[0]
        self.assertTrue(observed['passed'])
        self.assertIn('referenceExclusion', observed)


if __name__ == '__main__':
    unittest.main(verbosity=2)
