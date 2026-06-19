#!/usr/bin/env python3
"""ssbt — spreadsheet build tool.

Usage:
    ssbt build [--yml FILE] [--input FILE] [--output FILE] [--dry-run]
    ssbt test  [--yml FILE] [--input FILE]
    ssbt docs  [--yml FILE]

Models are defined in ssbt.yml with SQL transforms.
{{ ref('model_name') }} references another model's output as a table.
"""

import sys
from engine import main

if __name__ == "__main__":
    main()
