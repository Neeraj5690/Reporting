import os
from pathlib import Path
home = str(Path.home())
home = home.replace(os.sep, '/')
print(home)