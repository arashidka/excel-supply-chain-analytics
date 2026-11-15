import hashlib
import sys
from pathlib import Path

for p in map(Path, sys.argv[1:]):
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    print(f"{p.name}\t{p.stat().st_size}\t{h.hexdigest()}")
