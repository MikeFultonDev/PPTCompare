#!/bin/bash
# PowerPoint Comparison Tool Wrapper Script
# This script activates the virtual environment and runs the comparison tool
# Note: This script only works on macOS and Linux, not Windows

# Get the directory where the actual script is located (resolves symlinks)
SOURCE="${BASH_SOURCE[0]}"
while [ -h "$SOURCE" ]; do
  DIR="$( cd -P "$( dirname "$SOURCE" )" && pwd )"
  SOURCE="$(readlink "$SOURCE")"
  [[ $SOURCE != /* ]] && SOURCE="$DIR/$SOURCE"
done
SCRIPT_DIR="$( cd -P "$( dirname "$SOURCE" )" && pwd )"

# Activate the virtual environment
source "$SCRIPT_DIR/venv/bin/activate"

# Run the Python script with all passed arguments
python "$SCRIPT_DIR/ppt_compare.py" "$@"

# Made with Bob
