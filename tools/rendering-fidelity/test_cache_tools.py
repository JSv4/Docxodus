import copy
import importlib.util
from pathlib import Path
import unittest


def load(name, file):
    spec = importlib.util.spec_from_file_location(name, Path(__file__).with_name(file))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


canary = load('cache_canary', 'check-cached-results.py')
key = load('measurement_key', 'measurement-key.py')


def result():
    return {'candidate': 'abc', 'contractDigest': 'c1',
            'contract': {'referenceDigest': 'r1', 'environment': {'fonts': {'a.ttf': 'f1'}}},
            'summary': {'score': 87.19},
            'cases': [{'id': 'one', 'score': 99.5, 'adapter': {'pdfSha256': 'p1', 'executableSha256': 'e1'},
                       'pages': [{'ssim': 0.99, 'raster': {'candidateSha256': 'x1'}}]},
                      {'id': 'two', 'score': 80.0, 'adapter': {'pdfSha256': 'p2', 'executableSha256': 'e1'},
                       'pages': []}]}


class CachedResultCanary(unittest.TestCase):
    def test_identical_results_reproduce(self):
        self.assertEqual(canary.differences(result(), result()), [])

    def test_pdf_byte_hashes_are_per_run(self):
        live = result()
        live['cases'][0]['adapter']['pdfSha256'] = 'other'
        live['cases'][1]['adapter']['pdfSha256'] = 'other'
        self.assertEqual(canary.differences(result(), live), [])

    def test_any_measured_value_is_a_difference(self):
        for mutate, path in [
            (lambda r: r['cases'][0].update(score=99.4), ('cases', 0, 'score')),
            (lambda r: r['cases'][0]['pages'][0].update(ssim=0.98), ('cases', 0, 'pages', 0, 'ssim')),
            (lambda r: r['cases'][0]['pages'][0]['raster'].update(candidateSha256='x2'),
             ('cases', 0, 'pages', 0, 'raster', 'candidateSha256')),
            (lambda r: r['cases'][1]['adapter'].update(executableSha256='e2'), ('cases', 1, 'adapter', 'executableSha256')),
            (lambda r: r['contract']['environment']['fonts'].update({'a.ttf': 'f2'}),
             ('contract', 'environment', 'fonts', 'a.ttf')),
            (lambda r: r['summary'].update(score=87.2), ('summary', 'score')),
        ]:
            live = result()
            mutate(live)
            self.assertEqual(canary.differences(result(), live), [path])

    def test_missing_or_extra_fields_are_differences(self):
        live = result()
        del live['cases'][1]
        live['cases'][0]['pages'].append({'ssim': 1})
        self.assertIn(('cases', 1, 'id'), canary.differences(result(), live))
        self.assertIn(('cases', 0, 'pages', 1, 'ssim'), canary.differences(result(), live))

    def test_a_field_whose_cached_value_is_absent_differs_from_none(self):
        cached, live = result(), result()
        live['cases'][0]['note'] = None
        self.assertEqual(canary.differences(cached, live), [('cases', 0, 'note')])


class MeasurementKey(unittest.TestCase):
    components = {'libreofficeEnvironment': 'a', 'wordEnvironment': 'b', 'libreofficeReference': 'LO 25.8.7.3',
                  'image': ['ubuntu24', '20261102.1'], 'packages': ['qpdf=11.9.0-1']}

    def test_key_is_stable(self):
        self.assertEqual(key.compose(self.components), key.compose(copy.deepcopy(self.components)))

    def test_every_component_changes_the_key(self):
        for name in self.components:
            changed = copy.deepcopy(self.components)
            changed[name] = changed[name] + ['x'] if isinstance(changed[name], list) else changed[name] + 'x'
            self.assertNotEqual(key.compose(self.components), key.compose(changed), name)


if __name__ == '__main__':
    unittest.main()
