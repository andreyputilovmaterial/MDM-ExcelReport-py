"""Tests for column-header resolution (id vs human-readable) in the Excel report."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))

import util_format_mddread_inputs as uf


def _sample_inp():
    return {
        'report_type': 'TextFile',
        'source_file': '/tmp/a.txt',
        'report_datetime_utc': '2026-01-01T00:00:00Z',
        'report_datetime_local': '2026-01-01T00:00:00Z',
        'source_file_metadata': [],
        'report_scheme': {
            'columns': ['name', 'rawtextcontents'],
            'column_headers': {'name': 'Item name', 'rawtextcontents': 'File Contents'},
            'flags': [],
        },
        'sections': [
            {'name': 'filecontents', 'content': [{'name': '', 'rawtextcontents': 'hello'}]},
        ],
    }


class TestResolveColumnHeaders(unittest.TestCase):
    def test_text_mode_merges_global_and_section(self):
        result = uf.resolve_column_headers(
            ['name', 'label', 'extra'],
            {'name': 'Item name', 'label': 'Item label'},
            {'label': 'Section label'},
            'text',
        )
        self.assertEqual(result, {
            'name': 'Item name',
            'label': 'Section label',
            'extra': 'extra',
        })

    def test_text_mode_falls_back_to_id(self):
        self.assertEqual(uf.resolve_column_headers(['a'], {}, {}, 'text'), {'a': 'a'})

    def test_text_mode_preserves_empty_header(self):
        self.assertEqual(uf.resolve_column_headers(['a'], {'a': ''}, {}, 'text'), {'a': ''})

    def test_fieldname_mode_ignores_headers(self):
        result = uf.resolve_column_headers(
            ['name', 'label'],
            {'name': 'Item name', 'label': 'Item label'},
            {},
            'fieldname',
        )
        self.assertEqual(result, {'name': 'name', 'label': 'label'})

    def test_default_mode_is_text(self):
        self.assertEqual(uf.resolve_column_headers(['a'], {'a': 'A'}), {'a': 'A'})


class TestPrepDataframes(unittest.TestCase):
    def _sections_by_name(self, sections):
        return {s['name']: s for s in sections}

    def test_text_mode_uses_human_headers(self):
        config = {
            'columns': ['name', 'rawtextcontents'],
            'column_headers': {'name': 'Item name', 'rawtextcontents': 'File Contents'},
            'column_headers_mode': 'text',
        }
        sections = self._sections_by_name(uf.prep_dataframes(_sample_inp(), config))
        # The overview sheet intentionally blanks its own headers.
        self.assertEqual(sections['overview']['column_headers'], {'name': '', 'value': ''})
        self.assertEqual(sections['filecontents']['columns'], ['name', 'rawtextcontents'])
        self.assertEqual(sections['filecontents']['column_headers'], {
            'name': 'Item name',
            'rawtextcontents': 'File Contents',
        })

    def test_fieldname_mode_uses_fieldnames(self):
        config = {
            'columns': ['name', 'rawtextcontents'],
            'column_headers': {'name': 'Item name', 'rawtextcontents': 'File Contents'},
            'column_headers_mode': 'fieldname',
        }
        sections = self._sections_by_name(uf.prep_dataframes(_sample_inp(), config))
        self.assertEqual(sections['filecontents']['column_headers'], {
            'name': 'name',
            'rawtextcontents': 'rawtextcontents',
        })

    def test_default_mode_is_text(self):
        config = {
            'columns': ['name', 'rawtextcontents'],
            'column_headers': {'name': 'Item name', 'rawtextcontents': 'File Contents'},
        }
        sections = self._sections_by_name(uf.prep_dataframes(_sample_inp(), config))
        self.assertEqual(sections['filecontents']['column_headers'], {
            'name': 'Item name',
            'rawtextcontents': 'File Contents',
        })


class TestSanitizeHeader(unittest.TestCase):
    def test_plain_string(self):
        self.assertEqual(list(uf.sanitize('Item name')), [('Item name', '')])

    def test_rich_text_parts(self):
        value = {'parts': [{'text': 'A'}, {'text': 'B', 'role': 'added'}]}
        self.assertEqual(list(uf.sanitize(value)), [('A', ''), ('B', 'added')])


if __name__ == '__main__':
    unittest.main()
