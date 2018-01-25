User Manual for Summarize
=========================

Short instructions on using the summarize GUI tool.

Note: Summarize can be run from the command line as well, but that is out of scope for this manual. Minimal help is available running with "summarize --help".

Current limitations:

- All files to be summarized must be in a single directory (workaround: copy to same directory first)
- Rows are not sorted (workaround: sort in spreadsheet software)


Step one: Install
-----------------
- Download the latest release from: [dist/summarize.zip @ appveyor](https://ci.appveyor.com/project/alon/pump-summarize/build/artifacts)
- Unzip summarize.zip
- Locate summarize.exe inside the summarize directory
- For convenience: Create a shortcut by draging the exe to the desktop

Optional Step: Prepare summary.ini
----------------------------------
- In the directory containing the excel files you wish to summarize you can place a file to direct the summarizer.
- Example contents at the end of this file

Step two: Launch
----------------
- Double click the summarize.exe or shortcut you created

  ![After launch][launch]\


Step three: Drag excel files
----------------------------

- Open an explorer window with the excels you wish to summarize
- Drag in a single or multiple times the files you wish to summarize

  ![dragging file][drag]\


  - After the first drag a button appears
- Click the button

  ![clicking button][click]\


- Wait for the progress dialog to appear and progress.

  ![waiting for operation][progress]\


- Result file is named summary.xlsx unless a file already exists (from a previous invocation), in which case the first summary_N.xlsx available is used.
- The resulting file is opened automatically with the associated application (Microsoft Office Excel / Libreoffice Calc or otherwise).

  ![result spreadsheet][spreadsheet]\


- Enter pump_head value for each row (file)

  ![pump head entry][pump_head]\



Example optional summary.ini
============================
```ini
[global]
parameters=
[user_defined]
fields=
[half_cycle]
fields=Cruising Velocity [m/s],Cruising Flow Rate [LPM],Cruising Power In [W], Flow Rate [LPM],Cruising Power Out [W],Cruising Efficiency [%]
directions=down,up
```

[launch]: 01_after_launch.png
[drag]: 02_drag.png
[click]: 03_click.png
[progress]: 04_progress.png
[spreadsheet]: 05_spreadsheet.png
[pump_head]: 06_pump_head_entered.png
