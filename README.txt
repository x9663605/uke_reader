UKE Reader

QGIS plugin reading *.xlsx file and creating temporary vector layer (able to save).
Input is degree coordinates in 4326, output is decimal coordinates in 4326.

Instruction of installing new plugins in QGIS (Windows):
*	Paste extracted plugin in C:\Users\user\.qgis2\python\plugins\uke_reader
*	Turn on in QGIS Plugin Manager

Instruction of UKE Reader (sample data):
*	Open the plugin (Plugins -> UKE Reader)
*	Choose *.xlsx file in "Input"
*	Type the letter of column with X coordinate (latitude) in "X column".
*	Type the letter of column with Y coordinate (longitude) in "Y column".
*	Enter which row has first coordinates in "Start row".

In plugin has been used following modules:
*	qgis.core (QGIS support)
*	PyQt4 (QT Creator for GUI support)
*	openpyxl (excel files support)
*	re (regular expression)