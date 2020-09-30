Translation Analysis App
Creation of:
Michal Nahlik, School of Engineering, University of Warwick, 
michal.nahlik@warwick.ac.uk
Ewa Lelowicz-Nahlik, University of Portsmouth, 
ewa.lelowicz@gmail.com

This is a script designed to automate to an extent the process of analysing a large amount of translated text and offer a numerical score based on a human-defined collocation wordlist. This work is dedicated as assistance to the dissertation of Ewa Lelowicz-Nahlik of University of Portsmouth.
---------------------------------------------------------------------------------------------------
Files: 
Polish Translation Analysis Script - Demo - Visual Representation.mp4

Demo displaying the functionality as well as visualisation of workings of the app

Files:
working_wordlist_Historia.xlsx
working_wordlist_Sociology.xlsx
working_wordlist_Psychology.xlsx

Contain the complete wordlists which get updated by the script and work as the database for the scoring of the quality of the translation

Files:
Wordlist_Ela.xlsx
Wordlist_Magdalena.xlsx
Wordlist_Stanislaw.xlsx
Wordlist_Swietana.xlsx

Contain the wordlists for the individual samples. Please note that all the cells are collapsed for siplification of the script, however once opened with Microsoft Excel, one can freely expand the column size to see the full text within each of the cells.
----------------------------------------------------------------------------------------------------
NOTE:

For the following reasons, the script may prove to be troublesome to execute. Therefore it is advisable to contact the university's IS team for assistance in running the script.

The package contains the source code as well as an executable file for anyone unfamiliar with running a .py file via an Integrated development environment (IDE). However, since the resultant executable reads and writes data on local files (namely the excel spreadsheet of which data it analyzes), running the .exe file causes a false positive virus detection from the Windows Defender SmartScreen. It also happens since the executable is not verified and signed despite being from an acquianted source with which I have supported in creating the script.

To run the executable:
1. Extract 'Executable.zip' so that the 'Translation_Analysis_App.exe' is in the same directory as the excel spreadsheet wordlists.
2. Run Translation_Analysis.exe
3. If 'Windows protected your PC' warning is displayed:
	i. Press 'More info'
	ii. Press 'Run anyway'

Alternatively, a Python 3.8 IDE (such as PyCharm) can be installed along with the following packages: Tkinter, Openpyxl, pandas, docx2txt. Thereafter, within the same directory Translation_Analysis_App.py can be run.

There is also a short video contained within the directory, displaying a demo of the app in use along with a visualisation of the workings of the script.

