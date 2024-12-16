#!/usr/bin/env bash

DIRECTORY_LOG="$(pwd)/logs/"
DIRECTORY_SCH="~/Documents/Scheduled/"

setup() {
    echo "python3 DataPipeline.py --configure"
}

if [[ ! -d "$DIRECTORY_LOG" ]]; then
    echo "$(setup)"
fi

if [ ! -d "$DIRECTORY_LOG" ] && [ "$(ls -A "$DIRECTORY_SCH")" ]; then
    echo "python3 DataPipeline.py --task 4"
else
    echo "$DIRECTORY_SCH is empty. Add files then run again."
fi
