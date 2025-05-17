## 1.1) Creating Folders for Each Date in a Given Range


```python
import os
from datetime import date, timedelta

# Target path
base_path = r"E:\test_path"

# Define start and end dates
start_date = date(2025, 3, 1)
end_date = date(2025, 3, 10)

# Loop through each day and create a folder
current_date = start_date
while current_date <= end_date:
    folder_name = current_date.strftime("%Y-%m-%d")
    folder_path = os.path.join(base_path, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    current_date += timedelta(days=1)

print("Date Folders Created Successfully.")
```


    

## 1.2) Creating Folders Only for Working Days in a Date Range (Excluding Weekends)


```python
import os
from datetime import date, timedelta

# Target path
base_path = r"E:\test_path"

# Define start and end dates
start_date = date(2025, 3, 1)
end_date = date(2025, 3, 10)

# Loop through each day and create a folder for weekdays only
current_date = start_date
while current_date <= end_date:
    if current_date.weekday() < 5:  # 0 to 4 → Monday to Friday
        folder_name = current_date.strftime("%Y-%m-%d")
        folder_path = os.path.join(base_path, folder_name)
        os.makedirs(folder_path, exist_ok=True)
    current_date += timedelta(days=1)

print("Working Day Folders Created Successfully.")
```


    

## 2.1) Creating Folders with Date and Day Name for Each Day in a Date Range


```python
import os
from datetime import date, timedelta

# Target path
base_path = r"E:\test_path"

# Define start and end dates
start_date = date(2025, 3, 1)
end_date = date(2025, 3, 10)

# Loop through each day and create a folder
current_date = start_date
while current_date <= end_date:
    date_part = current_date.strftime("%Y-%m-%d")
    day_name = current_date.strftime("%A")  # Full day name (e.g., Monday)
    folder_name = f"{date_part}_{day_name}"
    folder_path = os.path.join(base_path, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    current_date += timedelta(days=1)

print("Date Folders with day names created successfully.")
```


    

## 2.2) Creating Folders for Working Days Only with Date and Day Name in Folder Name


```python
import os
from datetime import date, timedelta

# Target path
base_path = r"E:\test_path"

# Define start and end dates
start_date = date(2025, 3, 1)
end_date = date(2025, 3, 10)

# Loop through each day and create folders for weekdays only
current_date = start_date
while current_date <= end_date:
    if current_date.weekday() < 5:  # 0=Monday, ..., 4=Friday
        date_part = current_date.strftime("%Y-%m-%d")
        day_name = current_date.strftime("%A")  # Full day name
        folder_name = f"{date_part}_{day_name}"
        folder_path = os.path.join(base_path, folder_name)
        os.makedirs(folder_path, exist_ok=True)
    current_date += timedelta(days=1)

print("Working day folders with day names created successfully.")
```


    

## 3.1) Creating Date-Based Subfolders for Multiple Names


```python
import os
from datetime import date, timedelta

# Base path
base_path = r"E:\test_path"

# List of names
names = ['Umer', 'Ali', 'Ahmed']

# Date range
start_date = date(2025, 3, 1)
end_date = date(2025, 3, 10)

# Create folders
for name in names:
    name_folder_path = os.path.join(base_path, name)
    os.makedirs(name_folder_path, exist_ok=True)

    current_date = start_date
    while current_date <= end_date:
        folder_name = current_date.strftime("%Y-%m-%d")
        folder_path = os.path.join(name_folder_path, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        current_date += timedelta(days=1)

print("Date Folders Created Inside Each Name Folder Successfully.")
```


    

## 3.2) Creating Date-Based Subfolders for Each Name While Excluding Weekends


```python
import os
from datetime import date, timedelta

# Base path
base_path = r"E:\test_path"

# List of names
names = ['Umer', 'Ali', 'Ahmed']

# Date range
start_date = date(2025, 3, 1)
end_date = date(2025, 3, 10)

# Create folders
for name in names:
    name_folder_path = os.path.join(base_path, name)
    os.makedirs(name_folder_path, exist_ok=True)

    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() < 5:  # 0=Monday to 4=Friday
            folder_name = current_date.strftime("%Y-%m-%d")
            folder_path = os.path.join(name_folder_path, folder_name)
            os.makedirs(folder_path, exist_ok=True)
        current_date += timedelta(days=1)

print("Date Folders (Weekdays Only) Created Inside Each Name Folder Successfully.")

```


    

## 4.1) Creating Date and Day Name Subfolders for Multiple Names


```python
import os
from datetime import date, timedelta

# Base path
base_path = r"E:\test_path"

# List of names
names = ['Umer', 'Ali', 'Ahmed']

# Date range
start_date = date(2025, 3, 1)
end_date = date(2025, 3, 5)

# Create folders
for name in names:
    name_folder_path = os.path.join(base_path, name)
    os.makedirs(name_folder_path, exist_ok=True)

    current_date = start_date
    while current_date <= end_date:
        folder_name = current_date.strftime("%Y-%m-%d_%A")  # Add day name
        folder_path = os.path.join(name_folder_path, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        current_date += timedelta(days=1)

print("Date folders with day names created inside each name folder successfully.")

```

## 4.2) Creating Weekday Folders with Date and Day Name Inside Each Name Folder


```python
import os
from datetime import date, timedelta

# Base path
base_path = r"E:\test_path"

# List of names
names = ['Umer', 'Ali', 'Ahmed']

# Date range
start_date = date(2025, 3, 1)
end_date = date(2025, 3, 10)

# Create folders
for name in names:
    name_folder_path = os.path.join(base_path, name)
    os.makedirs(name_folder_path, exist_ok=True)

    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() < 5:  # 0=Monday, ..., 4=Friday
            folder_name = current_date.strftime("%Y-%m-%d_%A")  # Add day name
            folder_path = os.path.join(name_folder_path, folder_name)
            os.makedirs(folder_path, exist_ok=True)
        current_date += timedelta(days=1)

print("Weekday folders with date and day name created inside each name folder successfully.")
```
