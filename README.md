# Wavetec to Excel

## Run

Install Python 3.10 or newer, then run:

```powershell
python -m pip install -r requirements.txt
python app.py
```

Select one or more ZIP or TXT files, choose an `.xlsx` output path, and press **Convert (this may take a long time)**. In the file picker, hold `Ctrl` or `Shift` to select multiple input files. The application reads only `.txt` entries directly contained in each selected ZIP. Other entries, including nested ZIP files, are ignored. After parsing, select the dates to include in the report. Parsing and Excel generation run in background threads while a percentage progress dialog keeps the GUI responsive.

The first worksheet is `Report`. It contains four enlarged line charts stacked vertically: daily min/max speed, daily reduction, hourly min/max speed, and hourly reduction. Speed charts use 20 km/h y-axis steps; average reduction charts use 5 km/h steps and scale only to the observed averaged values. Charts use date-only or time-only x-axis categories, at most eight x-axis entries, actual measured first/last points, and gridlines. The heatmaps are separated into `Speed heatmap by day` and `Speed heatmap by hour`. Each groups first-measurement speeds into 5 km/h bins, displays absolute driver counts, and shades cells by the bin's percentage of drivers within the corresponding day/hour column. Only the first measurement of each vehicle is included. The visible report tables use columns A:G and are stacked below the charts, making the sheet suitable for portrait/vertical printing. They contain overall minimum/maximum speeds with their recorded timestamps, daily minimum/maximum speeds with timestamps, and hourly minimum/maximum speeds and reductions across all selected dates. The chart-source dates are displayed as readable date/time text. All charts are based only on the selected dates. The raw measurement rows are stored afterward in `Measurements 1`, `Measurements 2`, etc. Each measurement sheet contains Date, Time, Vehicle, First speed, Second speed, and Reduction columns with filterable headers. The first header row repeats on every printed page. If the selected input exceeds Excel's worksheet row limit, additional measurement sheets are created automatically.

## Build a Windows executable

```powershell
python -m pip install pyinstaller
pyinstaller --onefile --windowed --name WavetecToExcel app.py
```

The distributable executable is created in `dist/WavetecToExcel.exe`.