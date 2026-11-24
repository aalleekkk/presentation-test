#!/bin/bash
# usage: RAW=<dir> ./build.sh
set -e
cd "$(dirname "$0")"
export PPTX_SKILL=${PPTX_SKILL:-/root/.claude/skills/synced/c5c058a7-8513-4bb0-b0b9-25368f1033ca_07b7f7d3-4031-4e50-93fe-5619989faab3/pptx}
export NODE_PATH=$PWD/node_modules
node deck.js
python3 tools/transitions.py ../Sta_cini_nesto_lepim.pptx transitions.json
python3 $PPTX_SKILL/scripts/office/validate.py ../Sta_cini_nesto_lepim.pptx
