import copy
import importlib.util
from pathlib import Path
import unittest

from bench.common import identity
from bench.metrics import aggregate, case_score, CONTRACT

spec = importlib.util.spec_from_file_location('provider_gate', Path(__file__).with_name('check-provider-regressions.py'))
gate = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gate)


def case(cid, score, raster):
    return {'id': cid, 'category': 'paragraphs', 'status': 'measured',
            'checks': [{'name': 'full-visible-inventory', 'passed': True}],
            'score': score, 'close': False,
            'pages': [{'score': score, 'ssim': score / 100, 'inkF1': 1,
                       'precision': 1, 'recall': 1, 'raster': {'candidateSha256': raster}}]}


def result(track, cases):
    contract = {'metric': CONTRACT, 'referenceDigest': track, 'layoutProfile': 'default'}
    ids = [c['id'] for c in cases]
    return {'schema': 1, 'track': track, 'contract': contract, 'contractDigest': identity(contract),
            'suite': {'cases': ids}, 'cases': cases, 'summary': aggregate(cases, ids)}


class SeparateProviderGate(unittest.TestCase):
    def setUp(self):
        self.lo_before = result('libreoffice', [case('reviewed', 99, 'old'), case('other', 99, 'untouched')])
        self.lo_after = result('libreoffice', [case('reviewed', 90, 'new'), case('other', 99, 'untouched')])
        self.lo_main = copy.deepcopy(self.lo_before)
        self.word_before = result('word-online', [case('reviewed-word', 85, 'old')])
        self.word_after = result('word-online', [case('reviewed-word', 96, 'new')])
        self.policy = {'schema': 1, 'referenceTrack': 'word-online', 'layoutProfile': 'default',
                       'minimumWordScore': 95, 'cases': [{'originalCase': 'reviewed', 'wordCase': 'reviewed-word'}]}

    def decide(self):
        # Recompute valid result summaries after independent fault injection.
        for r in (self.lo_before, self.lo_after, self.lo_main, self.word_before, self.word_after):
            for c in r['cases']:
                c['score'] = case_score(c['pages'], c['checks'])
            r['summary'] = aggregate(r['cases'], r['suite']['cases'])
        return gate.decide(self.lo_before, self.lo_after, self.lo_main, self.word_before, self.word_after, self.policy)

    def test_word_improvement_is_adjudicated_and_lo_decline_stays_visible(self):
        report = self.decide()
        self.assertTrue(report['passed'])
        self.assertEqual(report['wordAuthoritativeCases'], ['reviewed'])
        self.assertTrue(report['referenceTracks']['libreoffice']['allFindings'])
        self.assertTrue(report['reviewedLibreOfficeProviderDifferences']['reviewed'])

    def test_unreviewed_libreoffice_decline_blocks(self):
        self.lo_after['cases'][1]['pages'][0]['precision'] = .999
        self.assertFalse(self.decide()['passed'])

    def test_word_metric_regression_blocks_even_when_overall_score_improves(self):
        self.word_after['cases'][0]['pages'][0]['precision'] = .999
        report = self.decide()
        self.assertFalse(report['passed'])
        self.assertEqual(report['wordAuthoritativeCases'], [])

    def test_word_mandatory_failure_cannot_adjudicate_lo(self):
        self.word_after['cases'][0]['checks'][0]['passed'] = False
        report = self.decide()
        self.assertFalse(report['passed'])
        self.assertEqual(report['wordAuthoritativeCases'], [])

    def test_mapped_original_still_requires_all_mandatory_checks(self):
        self.lo_after['cases'][0]['checks'][0]['passed'] = False
        self.assertFalse(self.decide()['passed'])

    def test_changed_main_or_candidate_proxy_pixels_block(self):
        for name in ('word_before', 'word_after'):
            with self.subTest(renderer=name):
                r = getattr(self, name)
                saved = r['cases'][0]['pages'][0]['raster']['candidateSha256']
                r['cases'][0]['pages'][0]['raster']['candidateSha256'] = 'different'
                self.assertFalse(self.decide()['passed'])
                r['cases'][0]['pages'][0]['raster']['candidateSha256'] = saved

    def test_missing_raster_hash_cannot_prove_equivalence(self):
        self.word_after['cases'][0]['pages'][0].pop('raster')
        self.assertFalse(self.decide()['passed'])

    def test_original_word_score_target_is_required(self):
        self.word_after['cases'][0]['pages'][0]['score'] = 94.99
        self.assertFalse(self.decide()['passed'])

    def test_word_profile_or_provider_cannot_change(self):
        self.word_after['track'] = 'libreoffice'
        with self.assertRaisesRegex(ValueError, 'registered Word'):
            self.decide()
        self.word_after['track'] = 'word-online'
        self.word_after['contract']['layoutProfile'] = 'libreoffice'
        self.word_after['contractDigest'] = identity(self.word_after['contract'])
        with self.assertRaisesRegex(ValueError, 'explicit renderer layout'):
            self.decide()

    def test_new_reference_cannot_replace_the_accepted_reference(self):
        self.word_after['contract']['referenceDigest'] = 'new reference'
        self.word_after['contractDigest'] = identity(self.word_after['contract'])
        with self.assertRaisesRegex(ValueError, 'Measurement contract changed'):
            self.decide()

    def test_missing_word_case_cannot_reduce_the_protected_corpus(self):
        self.policy['cases'].append({'originalCase': 'other', 'wordCase': 'not-measured'})
        with self.assertRaisesRegex(ValueError, 'exact registered case coverage'):
            self.decide()

    def test_duplicate_mapping_cannot_expand_authority(self):
        self.policy['cases'].append(dict(self.policy['cases'][0]))
        with self.assertRaisesRegex(ValueError, 'Duplicate'):
            self.decide()

    def test_missing_baseline_metric_cannot_disable_that_regression_check(self):
        self.word_before['cases'][0]['pages'][0].pop('precision')
        self.assertFalse(self.decide()['passed'])


if __name__ == '__main__':
    unittest.main()
