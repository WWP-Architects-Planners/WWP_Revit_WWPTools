import os
import sys

script_dir = os.path.dirname(__file__)
parent_dir = os.path.dirname(script_dir)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

from massidtool_core import publish_mass_level_metrics


if __name__ == "__main__":
    publish_mass_level_metrics(xaml_dir=script_dir)
