Build instructions for creating an executable of `src/main.py`

1. Open a PowerShell or CMD prompt.
2. Change to the project folder `sinapi-gui-app` (the batch is located here):

   cd "c:\Users\walyson.ferreira\Desktop\InterfacePython\sinapi-gui-app"

3. Run the build script:

   build_exe.bat

What the script does:
- Ensures PyInstaller is installed (via `pip install --upgrade pyinstaller`).
- Runs PyInstaller in `src` producing a single-file executable named `SinapiApp`.
- `LINKS_SICRO.txt` is included in the exe bundle as data so `sicro.resource_path` can find it in the onefile build.

Notes & troubleshooting:
- If you use a virtual environment, activate it before running the script.
- If PyInstaller is already installed on your system, the script will use it.
- The generated executable will be at `sinapi-gui-app/src/dist/SinapiApp.exe`.
- If the GUI doesn't start, run the executable from a console to see stderr output.

Optional improvements:
- Add other resource files (`resources/`, images) using `--add-data` if needed.
- Create an installer (Inno Setup / WiX) to package and sign the binary.
