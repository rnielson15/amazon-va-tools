#!/bin/bash
cd "$(dirname "$0")" || exit 1
if [ ! -d ".venv" ]; then
  echo ""
  echo "First-time setup — this happens only once and takes a minute or two..."
  python3 -m venv .venv
  source .venv/bin/activate
  pip install -r requirements.txt
else
  source .venv/bin/activate
fi
echo ""
echo "Starting the Bleeders Report app — your browser will open in a moment."
echo "Leave this window open while you use it."
streamlit run app.py
