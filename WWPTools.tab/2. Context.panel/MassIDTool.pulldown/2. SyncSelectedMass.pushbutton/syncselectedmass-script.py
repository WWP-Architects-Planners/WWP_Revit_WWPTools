import os
import sys

script_dir = os.path.dirname(__file__)
parent_dir = os.path.dirname(script_dir)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

from massidtool_core import run_sync_selected


if __name__ == "__main__":
    run_sync_selected()
