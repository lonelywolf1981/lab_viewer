README (Windows 11 test)

1) Extract to a simple path, e.g. C:\Tools\LeMuReViewer\lemure_viewer_win7
   Avoid Cyrillic and avoid Program Files.

2) Install Python.
   - If you want future Windows 7 compatibility: install Python 3.8.x
   - For Windows 11 only: Python 3.11/3.12 is fine
   Check: python -V (or py -V)

3) Run Setup_Once.cmd once.
   If it fails, open setup.log in the same folder.

4) Run Start_Viewer.cmd (or Start_Viewer_DEBUG.cmd).
   Open: http://127.0.0.1:8787/

If the plot is blank and you have no internet:
- Plotly is loaded from CDN in templates/index.html. Ask for offline build.
