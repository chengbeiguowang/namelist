打包EXE
```
pyinstaller -F -w main.py -p namelist.py  --hidden-import openpyxl --hidden-import os.path --hidden-import sys