#!/usr/bin/env sh
set -eu

PROJECT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$PROJECT_DIR"

python3 scripts/create_demo_data.py
python3 pile_photo_checker.py examples/demo_input --output examples/demo_output

printf '%s\n' "Demo complete. Open examples/demo_output/pile_photo_completeness.xlsx"
