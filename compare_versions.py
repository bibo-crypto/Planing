import os, hashlib, sys
root1 = r'E:\Planing'
root2 = r'E:\delta-converter'
files1 = {}
for dirpath, _, filenames in os.walk(root1):
    for f in filenames:
        if f.endswith(('.py','.json','.txt','.spec','.bat','.iss')):
            p = os.path.join(dirpath, f)
            rel = os.path.relpath(p, root1)
            if rel.startswith(('.git','build','dist','venv','__pycache__')):
                continue
            with open(p, 'rb') as fh:
                files1[rel] = hashlib.md5(fh.read()).hexdigest()
files2 = {}
for dirpath, _, filenames in os.walk(root2):
    for f in filenames:
        if f.endswith(('.py','.json','.txt','.spec','.bat','.iss')):
            p = os.path.join(dirpath, f)
            rel = os.path.relpath(p, root2)
            if rel.startswith(('.git','build','dist','venv','__pycache__')):
                continue
            with open(p, 'rb') as fh:
                files2[rel] = hashlib.md5(fh.read()).hexdigest()
for rel in sorted(set(files1) | set(files2)):
    if files1.get(rel) != files2.get(rel):
        print(rel)
