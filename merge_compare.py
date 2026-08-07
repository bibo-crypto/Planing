from pathlib import Path
import difflib
files = [
    'dfm_lookup.py','excel_exporter.py','gui.py','kamal_excel_exporter.py','kamal_parser.py','kamal_tab.py',
    'lotti_logic.py','magazino_logic.py','magazino_filato_tab.py','modern_widgets.py','ordine_kamal.py','ordini_elvy.py',
    'situazione_loaders.py','situazione_tab.py','situazione_settimana_logic.py','situazione_settimana_tab.py'
]
root1 = Path(r'E:\Planing')
root2 = Path(r'E:\delta-converter')
for rel in files:
    p1 = root1 / rel
    p2 = root2 / rel
    if not p1.exists() or not p2.exists():
        print(f'-- {rel}: missing')
        continue
    txt1 = p1.read_text(encoding='utf-8', errors='ignore').splitlines()
    txt2 = p2.read_text(encoding='utf-8', errors='ignore').splitlines()
    if txt1 == txt2:
        print(f'-- {rel}: identical')
        continue
    print(f'-- {rel}: differs')
    diff = difflib.unified_diff(txt1, txt2, fromfile='current', tofile='external', n=2)
    count = 0
    for line in diff:
        if line.startswith(('---','+++','@@')) or line.startswith(('+','-')):
            print(line)
            count += 1
        if count >= 120:
            break
    print('')
