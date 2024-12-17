#!/usr/bin/env bash

DIRECTORY_LOG=$(realpath "$(pwd)\logs" | tr '\\' '/')
DIRECTORY_SCH=$(realpath "$HOME\Documents\Scheduled" | tr '\\' '/')

setup() {
    if [[ ! -d "venv" ]]; then
      python -m venv venv
    fi
    source venv/Scripts/activate
    pip install -r requirements.txt
    python3 DataPipeline.py --configure
}

if [[ ! -d "$DIRECTORY_LOG" ]]; then
    setup
fi

if [[ -d "$DIRECTORY_SCH" && "$(find "$DIRECTORY_SCH" -mindepth 1 | head -n 1)" ]]; then
    if [[ -d "venv" ]]; then
      source venv/Scripts/activate
    else
      setup
    fi
    pip install -r requirements.txt
    python DataPipeline.py --task 4
else
    echo "$DIRECTORY_SCH is either does not exist or is empty. Create the Directory and/or add files then run again."
fi
