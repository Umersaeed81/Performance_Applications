```python
# =====📦 Import Required Libraries =====
import os  # 📁 Operating system interface
import pptx  # 📊 python-pptx (PowerPoint automation, if needed)
import openpyxl  # 📘 For creating and manipulating Excel workbooks
import win32com.client  # 🪟 For automating PowerPoint via COM
from datetime import datetime, time  # 🕒 For timestamping filenames and handling cell values
```


```python
# =====📂 Define File Paths and Naming Convention =====

# 📄 Path to the PowerPoint file we want to read
ppt_path = r"D:/Advance_Data_Sets/UMTS_SunSet/PPT/PPT_Temp/Center_Region_04062025.pptx"

# 📁 Directory where the new Excel file will be saved
directory = 'D:/Advance_Data_Sets/UMTS_SunSet/PPT/PPT_Temp'
base_name = 'graphs_data'
extension = 'xlsx'

# 🕒 Generate a timestamped filename (e.g., graphs_data_20250604_142530.xlsx)
now = datetime.now().strftime('%Y%m%d_%H%M%S')
excel_path = f"{directory}/{base_name}_{now}.{extension}"
```


```python
# =====📽️ Open PowerPoint Presentation via COM =====

# 🚀 Launch PowerPoint (hidden window) and open the target presentation
ppt_app = win32com.client.Dispatch("PowerPoint.Application")
ppt_app.Visible = True  # 👁️ Make PowerPoint visible; can set to False if running in background

# 📂 Open the PowerPoint file
presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=False)
```


```python
# =====📊 Initialize a New Excel Workbook =====

# 🧠 Create a workbook in memory
wb = openpyxl.Workbook()

# 🗑️ By default, openpyxl creates a sheet named "Sheet"; remove it so we can add custom-named sheets
wb.remove(wb.active)
```


```python
# =====📈 Function to Extract Chart Data from PowerPoint =====
def extract_chart_data(chart):
    """
    🧩 Extracts the raw data from a chart's embedded Excel worksheet.

    Returns:
        📦 A tuple of tuples representing all used cells in the chart's data sheet.
    """
    chart_data = chart.ChartData
    chart_workbook = chart_data.Workbook        # 📘 COM object for the embedded Excel workbook
    chart_sheet = chart_workbook.Worksheets(1)  # 📄 The first worksheet in the embedded workbook
    used_range = chart_sheet.UsedRange          # 🔍 All cells that contain data
    return used_range.Value                     # 📊 Value is a tuple of tuples
```


```python
# =====📤 Function to Write Chart Data into an Excel Worksheet =====
def write_chart_data_to_sheet(ws, data):
    """
    ✍️ Writes a 2D tuple (rows × columns) of chart data into the given openpyxl worksheet.
    🕒 Converts any timezone-aware datetime/time to naive before writing.
    """
    row = 1
    for r in data:
        for col_idx, cell_value in enumerate(r, start=1):
            # 🧹 Remove timezone info if present (openpyxl does not support tz-aware)
            if isinstance(cell_value, datetime) and cell_value.tzinfo is not None:
                cell_value = cell_value.replace(tzinfo=None)
            elif isinstance(cell_value, time) and cell_value.tzinfo is not None:
                cell_value = cell_value.replace(tzinfo=None)
            ws.cell(row=row, column=col_idx, value=cell_value)  # 🧾 Write value to worksheet
        row += 1  # ➕ Move to next row
```


```python
# =====📥 Function to Export All Charts from Slides to Excel Sheets =====
def export_charts_to_excel(presentation, wb):
    """
    🔄 Iterates through every slide in the PowerPoint presentation, finds each chart-shaped shape,
    📊 extracts its data, and writes that data into a new worksheet in the provided Excel workbook.
    """
    for slide_index, slide in enumerate(presentation.Slides, start=1):
        chart_count = 0  # 🧮 Counter for charts on the current slide

        for shape in slide.Shapes:
            # 🔍 Check if the shape is a chart
            if shape.HasChart:
                chart_count += 1  # ➕ Increment chart counter
                sheet_name = f"Slide_{slide_index}_Chart_{chart_count}"
                # 🆕 Create a new worksheet named after slide and chart number
                ws = wb.create_sheet(title=sheet_name)

                # 📥 Extract data from the PowerPoint chart and write into the worksheet
                chart_data = extract_chart_data(shape.Chart)
                write_chart_data_to_sheet(ws, chart_data)

```


```python
# =====✅ Execute Export Process and Cleanup =====

# 📤 Extract every chart into its own sheet in the workbook
export_charts_to_excel(presentation, wb)

# 💾 Save the Excel workbook to disk at the timestamped path
wb.save(excel_path)

# ❌📽️ Close the PowerPoint presentation and quit the application
presentation.Close()
ppt_app.Quit()

# ✅ Confirmation message
print(f"📁 Chart data exported successfully to: {excel_path}")
```

    Chart data exported successfully to: D:/Advance_Data_Sets/UMTS_SunSet/PPT/PPT_Temp/graphs_data_20250604_115932.xlsx
    
