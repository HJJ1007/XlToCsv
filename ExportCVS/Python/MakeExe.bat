pip install pyinstaller
pip install --upgrade pyinstaller
pyinstaller -w -F -n "xlToCsv" --icon="image/app.ico" --add-data="image/*;image" --additional-hooks-dir=. app.py 

MOVE .\dist\xlToCsv.exe .\xlToCsv.exe

@RD /S /Q .\build
@RD /S /Q .\dist
DEL /S /F /Q .\xlToCsv.spec
PAUSE