pip install pyinstaller
python ico2base64.py
pyinstaller --onefile --windowed --icon="icon.ico" --name="pdf-tool" main.py