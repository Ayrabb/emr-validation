# run.py  — PyInstaller entry point
import sys
import os
import argparse

# Fix imports when frozen (PyInstaller one-dir build)
if getattr(sys, 'frozen', False):
    sys.path.insert(0, sys._MEIPASS)
    # Place reports/ next to the .exe, not inside _internal/
    os.environ.setdefault('REPORT_DIR', os.path.join(os.path.dirname(sys.executable), 'reports'))

import uvicorn
from app.main import app

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8765)
    parser.add_argument('--host', type=str, default='127.0.0.1')
    args = parser.parse_args()

    print(f'[RADET Engine] Starting on http://{args.host}:{args.port}')
    uvicorn.run(app, host=args.host, port=args.port, reload=False)