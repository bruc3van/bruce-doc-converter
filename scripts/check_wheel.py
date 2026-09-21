"""Ensure wheels include every format module and locked Node runtime source."""
from pathlib import Path
from zipfile import ZipFile

wheel = max(Path('dist').glob('*.whl'), key=lambda path: path.stat().st_mtime)
with ZipFile(wheel) as archive:
    names = set(archive.namelist())
    for relative in ['output.py', 'formats/common.py', 'formats/images.py', 'formats/docx.py',
                     'formats/xlsx.py', 'formats/pptx.py', 'formats/pdf.py',
                     'md_to_docx/markdown-converter.js', 'md_to_docx/package-lock.json']:
        assert 'bruce_doc_converter/' + relative in names, relative
    assert not any('/node_modules/' in name for name in names), 'wheel contains node_modules'
print('Verified wheel contents:', wheel.name)
