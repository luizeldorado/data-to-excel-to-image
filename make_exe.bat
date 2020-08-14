pyinstaller ^
--onefile ^
--windowed ^
--add-data "qtui_main.ui;." ^
--add-data "translations/*.qm;translations" ^
gui.py

pause