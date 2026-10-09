import logging
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import fitz

from src.pdf_compiler import PDFCompiler
from src.pdf_util import PDFUtility


def test_filename_mode_merges_in_numeric_order_and_excludes_output(tmp_path):
    for name in ('10.pdf', '2.PDF', '1.pdf', 'combined.pdf'):
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((50, 50), name)
            if name == '2.PDF':
                page = doc.new_page()
                page.insert_text((50, 50), 'second page')
            doc.save(tmp_path / name)
    (tmp_path / 'unused.txt').write_text('Ignore non-document files.')
    (tmp_path / 'nested').mkdir()
    with fitz.open() as doc:
        doc.new_page()
        doc.save(tmp_path / 'nested' / 'hidden.pdf')

    gui = SimpleNamespace(
        entry_var1=Mock(get=Mock(return_value=str(tmp_path))),
        toc_var=Mock(get=Mock(return_value='pdf_names')),
        get_output_as=lambda: 'combined.pdf',
        pas_check_var=Mock(get=Mock(return_value=False)),
        entry_var5=Mock(),
        pb1={'value': 0},
        root=Mock(),
        btn_go=Mock(),
        logger=logging.getLogger('test_filename_merge'),
    )
    class Progress(dict):
        def place(self, **kwargs):
            pass
    gui.pb1 = Progress(value=0)
    util = PDFUtility(gui)
    util.convert_to_pdf = Mock(side_effect=AssertionError('Unexpected RTF conversion'))
    compiler = PDFCompiler(gui, util)
    compiler.combine_pdfs()
    compiler.combine_pdfs()  # Re-running must not append the previous result.

    with fitz.open(tmp_path / 'combined.pdf') as merged:
        assert 'Table of Contents' in merged[0].get_text()
        assert [page.get_text().strip() for page in list(merged)[1:]] == [
            '1.pdf', '2.PDF', 'second page', '10.pdf'
        ]
        assert merged.get_toc() == [[1, 'Table of Contents', 1], [1, '1.pdf', 2], [1, '2.PDF', 3], [1, '10.pdf', 5]]
        assert [link['page'] for link in merged[0].get_links()] == [1, 2, 4]
    util.convert_to_pdf.assert_not_called()


def test_filename_mode_converts_rtf_only_and_mixed_inputs(tmp_path):
    class Progress(dict):
        def place(self, **kwargs):
            pass
    gui = SimpleNamespace(
        entry_var1=Mock(get=Mock(return_value=str(tmp_path))),
        toc_var=Mock(get=Mock(return_value='pdf_names')),
        get_output_as=lambda: 'combined.pdf',
        pas_check_var=Mock(get=Mock(return_value=False)),
        entry_var5=Mock(), pb1=Progress(value=0), root=Mock(), btn_go=Mock(), logger=Mock(),
    )
    for name in ('2.RTF', '10.rtf'):
        (tmp_path / name).write_text(r'{\rtf1\ansi Test}')
    util = PDFUtility(gui)
    def convert(source, folder):
        output = str(Path(folder) / (Path(source).stem + '.pdf'))
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((50, 50), 'converted ' + Path(source).name)
            doc.save(output)
        return output
    util.convert_metadata_rtf = Mock(side_effect=convert)
    util.convert_to_pdf = Mock(side_effect=AssertionError('Do not use legacy Word conversion'))
    compiler = PDFCompiler(gui, util)
    compiler.combine_pdfs()
    with fitz.open(tmp_path / 'combined.pdf') as doc:
        assert doc.get_toc() == [[1, 'Table of Contents', 1], [1, '2.pdf', 2], [1, '10.pdf', 3]]
        assert [link['page'] for link in doc[0].get_links()] == [1, 2]
    for name in ('1.pdf', '2.PDF'):
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((50, 50), name)
            doc.save(tmp_path / name)
    compiler.combine_pdfs()
    with fitz.open(tmp_path / 'combined.pdf') as doc:
        assert len(doc) == 4  # Same-name RTF/PDF is included once; previous output excluded.
        assert [row[1] for row in doc.get_toc()] == ['Table of Contents', '1.pdf', '2.pdf', '10.pdf']
        assert doc[2].get_text().strip() == 'converted 2.RTF'
    assert util.convert_metadata_rtf.call_count == 4  # Both RTFs refreshed on each run.
    util.convert_to_pdf.assert_not_called()


def test_filename_bookmarks_support_chinese_and_leave_simple_mode_unchanged(tmp_path):
    source = tmp_path / '报告1.pdf'
    with fitz.open() as doc:
        doc.new_page()
        doc.save(source)
    gui = SimpleNamespace(entry_var1=Mock(get=Mock(return_value=str(tmp_path))), logger=Mock())
    util = PDFUtility(gui)
    for enabled in (True, False):
        output = tmp_path / ('bookmarked.pdf' if enabled else 'simple.pdf')
        assert util.combine_pdfs_simple([str(source)], str(output), filename_bookmarks=enabled)
        with fitz.open(output) as doc:
            assert doc.get_toc() == ([[1, '报告1.pdf', 1]] if enabled else [])
