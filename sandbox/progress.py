import json
import os
from pathlib import Path
import sys
import time
path = Path('/var/lib/a0-office/progress.json')
path.parent.mkdir(parents=True, exist_ok=True)
temporary = path.with_suffix('.tmp')
temporary.write_text(json.dumps({'boot':os.environ.get('A0_OFFICE_BOOT'), 'state': sys.argv[1], 'message': sys.argv[2], 'updated_at': time.time()}))
temporary.replace(path)
