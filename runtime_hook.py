"""Make bundled NumPy/Pandas native DLLs discoverable on Windows."""

from __future__ import annotations

import os
import sys


if sys.platform == "win32":
    bundle_root = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    dll_handles = []
    for directory_name in ("numpy.libs", "pandas.libs"):
        directory = os.path.join(bundle_root, directory_name)
        if os.path.isdir(directory):
            dll_handles.append(os.add_dll_directory(directory))
            os.environ["PATH"] = directory + os.pathsep + os.environ.get("PATH", "")
