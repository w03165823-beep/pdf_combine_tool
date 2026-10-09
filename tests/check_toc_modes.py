"""Run directly without pytest: python tests/check_toc_modes.py."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import csv
import tempfile
import fitz
from src.toc_merge import metadata_entries, merge_with_toc
import runpy
from types import SimpleNamespace
from unittest.mock import Mock, patch
from src.pdf_compiler import PDFCompiler
from src.pdf_util import PDFUtility


def check():
    old_tests = runpy.run_path(str(Path(__file__).with_name('test_filename_merge.py')))
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        for name, test in old_tests.items():
            if name.startswith('test_'):
                folder = root / name
                folder.mkdir()
                test(folder)
        for name in ('first.pdf', 'second.pdf'):
            with fitz.open() as doc:
                page = doc.new_page()
                page.insert_text((50, 50), name)
                doc.save(root / name)
        metadata = root / 'metadata.csv'
        def write_rows(rows):
            with metadata.open('w', encoding='utf-8-sig', newline='') as handle:
                writer = csv.writer(handle)
                writer.writerow(['Order', 'TFL', 'FileName', 'Title'])
                writer.writerows(rows)
        rows = [[10, 'T', 'first.pdf', '表：中文标题'], [2, 'F', 'second.pdf', '图：直接使用标题']]
        write_rows(rows)
        output = root / 'output.pdf'
        entries = metadata_entries(str(metadata), str(root), str(output))
        assert [Path(p).name for p, _ in entries] == ['second.pdf', 'first.pdf']
        gui = SimpleNamespace(entry_var1=Mock(get=Mock(return_value=str(root))),
                              entry_var2=Mock(get=Mock(return_value=str(metadata))),
                              toc_var=Mock(get=Mock(return_value='use_meta')),
                              get_output_as=lambda: 'output.pdf',
                              pas_check_var=Mock(get=Mock(return_value=False)),
                              entry_var5=Mock(), btn_go=Mock(), logger=Mock())
        PDFCompiler(gui, PDFUtility(gui)).combine_pdfs()
        with fitz.open(output) as doc:
            assert doc.get_toc() == [[1, 'Table of Contents', 1], [1, rows[1][3], 2], [1, rows[0][3], 3]]
            assert [link['page'] for link in doc[0].get_links()] == [1, 2]
            assert doc[1].get_text().strip() == 'second.pdf'
        import pandas as pd
        xlsx = root / 'metadata.xlsx'
        pd.read_csv(metadata).to_excel(xlsx, index=False)
        assert metadata_entries(str(xlsx), str(root), str(output)) == entries
        rtf_name = 'f-15-01-03-01-tte-forest'
        (root / (rtf_name + '.rtf')).write_text(r'{\rtf1\ansi Forest test}')
        write_rows([[2, 'F', rtf_name.upper(), 'Forest title'], [1, 'T', 'SECOND.PDF', 'PDF title']])
        matched = metadata_entries(str(metadata), str(root), str(output))
        assert Path(matched[1][0]).name == rtf_name + '.rtf'
        for spelling in (rtf_name.upper() + '.RTF', rtf_name.title(), '  ' + rtf_name.upper() + '  '):
            write_rows([[2, 'F', spelling, 'Forest title'], [1, 'T', 'SECOND.PDF', 'PDF title']])
            assert metadata_entries(str(metadata), str(root), str(output)) == matched
        (root / 'UPPER-SOURCE.RTF').write_text(r'{\rtf1\ansi Upper test}')
        write_rows([[1, 'T', 'upper-source', 'Upper title']])
        assert Path(metadata_entries(str(metadata), str(root), str(output))[0][0]).name == 'UPPER-SOURCE.RTF'
        write_rows([[2, 'F', rtf_name.upper(), 'Forest title'], [1, 'T', 'SECOND.PDF', 'PDF title']])
        word = Mock()
        def export(destination, format_code):
            assert format_code == 17
            with fitz.open() as document:
                page = document.new_page()
                page.insert_text((50, 50), 'Converted RTF')
                document.save(destination)
        word.Documents.Open.return_value.ExportAsFixedFormat.side_effect = export
        with patch('src.pdf_util.win32com.client.DispatchEx', return_value=word) as dispatch:
            compiler = PDFCompiler(gui, PDFUtility(gui))
            compiler.combine_pdfs()
            compiler.combine_pdfs()
            assert dispatch.call_count == 2  # Refresh the cached PDF.
            assert word.Quit.call_count == 2
            assert word.Documents.Open.call_args.args[0] == str(root / (rtf_name + '.rtf'))
        assert (root / '_PDF' / (rtf_name + '.pdf')).exists()
        with fitz.open(output) as doc:
            assert [row[1] for row in doc.get_toc()] == ['Table of Contents', 'PDF title', 'Forest title']
            assert doc[2].get_text().strip() == 'Converted RTF'
        with patch('src.pdf_util.win32com.client.DispatchEx', side_effect=RuntimeError('Word unavailable')):
            try:
                PDFCompiler(gui, PDFUtility(gui)).combine_pdfs()
            except RuntimeError as error:
                assert 'RTF conversion failed' in str(error)
            else:
                raise AssertionError('Conversion failure was ignored')
        with fitz.open(output) as doc:
            assert doc[2].get_text().strip() == 'Converted RTF'
        for invalid in (
            [[1, 'T', 'first.pdf', 'a'], [1, 'L', 'second.pdf', 'b']],
            [[1, 'X', 'first.pdf', 'a']],
            [[1, 'T', 'missing.pdf', 'a']],
            [[1, 'T', 'first.pdf', '']],
            [[1, 'T', '../first.pdf', 'a']],
        ):
            write_rows(invalid)
            try:
                metadata_entries(str(metadata), str(root), str(output))
            except ValueError:
                pass
            else:
                raise AssertionError(f'Accepted invalid metadata: {invalid}')
        # A long Unicode title spans TOC pages; every wrapped line must target its source.
        long_entries = [(str(root / 'first.pdf'), '中文标题与长标题测试' * 260),
                        (str(root / 'second.pdf'), 'Final item')]
        count = merge_with_toc(long_entries, str(output))
        assert count > 1
        with fitz.open(output) as doc:
            assert doc.get_toc()[1][2] == count + 1
            assert doc.get_toc()[2][2] == count + 2
            links = [link for page in list(doc)[:count] for link in page.get_links()]
            assert all(link['page'] == count for link in links[:-1])
            assert links[-1]['page'] == count + 1
        merge_with_toc(entries, str(output), password='owner-test')
        with fitz.open(output) as doc:
            assert doc.authenticate('owner-test')
            assert len(doc[0].get_links()) == 2
        # Retain a compact visual QA sample.
        preview = Path('tmp/pdfs')
        preview.mkdir(parents=True, exist_ok=True)
        merge_with_toc([(str(root / 'first.pdf'), '表1：受试者基本信息 - Baseline information'),
                        (str(root / 'second.pdf'), '图1：变化趋势 ' * 10)], str(preview / 'toc_preview.pdf'))
    print('PASS: both modes, bookmarks, clickable TOC, Chinese, wrapping, multi-page offsets, validation and encryption')


if __name__ == '__main__':
    check()
