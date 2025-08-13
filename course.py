import pandas as pd
import re
file="Time-Table, FSC, Fall-2025.xlsx"

sheet_name=pd.ExcelFile(file).sheet_names
print(sheet_name)

time_table_data={}

for names in sheet_name:
    if names!="Welcome":
        data=pd.read_excel(file,sheet_name=names)
        time_table_data[names]=data

#print(time_table_data)

def reshape_timetable(df, day_name):
    # Drop empty rows & columns
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    # Find "Room" header row
    header_row_index = df.index[df.iloc[:, 0].astype(str).str.contains("Room", case=False, na=False)]
    if len(header_row_index) == 0:
        return pd.DataFrame(columns=["Day", "Course Name", "Class Time", "Room No", "Batch"])
    
    header_row = header_row_index[0]

    # Batch row is right after header row in the "BS AI 2025" style
    batch_row = header_row - 1 if header_row > 0 else None

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    # Identify time columns and their batch names
    time_pattern = re.compile(r"\d{2}:\d{2}-\d{2}:\d{2}")
    time_columns = [col for col in df.columns if isinstance(col, str) and time_pattern.search(col)]

    # Extract batch names from the original Excel (if available)
    batch_names_map = {}
    if batch_row is not None and batch_row >= 0:
        original = pd.read_excel(file, sheet_name=day_name)
        original = original.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
        header_row_index_orig = original.index[original.iloc[:, 0].astype(str).str.contains("Room", case=False, na=False)][0]
        batch_row_data = original.iloc[batch_row]
        for col in time_columns:
            batch_names_map[col] = batch_row_data.get(col, None)

    results = []
    for _, row in df.iterrows():
        room = row.get("Room", "Unknown")
        for col in time_columns:
            raw_course = str(row[col]) if pd.notna(row[col]) else "Free Slot"

            # Extract actual time if embedded in course name
            found_time = time_pattern.search(raw_course)
            if found_time:
                actual_time = found_time.group()
                course_name = raw_course.replace(actual_time, "").strip()
            else:
                actual_time = col
                course_name = raw_course

            # Get batch name from mapping
            batch_name = batch_names_map.get(col, None)

            results.append({
                "Day": day_name,
                "Course Name": course_name,
                "Class Time": actual_time,
                "Room No": room,
                "Batch": batch_name
            })
    
    return pd.DataFrame(results)

# Process all days
event_tables = {day: reshape_timetable(time_table_data[day], day) for day in time_table_data}

# Example: Monday
print(event_tables["Monday"].head(30))