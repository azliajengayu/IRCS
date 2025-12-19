#!/bin/bash

SCRIPT_DIR = "$( cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

cd /d/python
ls dist13
./dit13/bootstrap_env.exe
source .venv/Scripts/activate

python -m pip install --no-index IRCS/modules13/pywin32-311-cp313-cp313-win_amd64.whl
python -m pip install --no-index IRCS/modules13/xlwings-0.33.16-cp313-cp313-win_amd64.whl

echo "Setup complete!"