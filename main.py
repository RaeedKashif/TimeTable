from flask import Flask, render_template, request, jsonify
import pandas as pd
from datetime import datetime
from TimeTable import get_time_table  # Import your existing script
import re

app = Flask(__name__)

# Load and preprocess the timetable data
def preprocess_timetable():
    df = get_time_table()
    
    # 1. Separate theory and lab classes
    df['Type'] = df.apply(lambda row: 'Lab' if 'Lab' in str(row['Course Name']) else row['Type'], axis=1)
    
    # Replace this section normalization:
# df['Section'] = df['Section'].apply(lambda x: x.split('-')[0] + '-' + x.split('-')[1][0] 
#                                   if x and '-' in str(x) and len(str(x).split('-')[1]) > 1 else x)

    # With this version that preserves A1, A2, etc. patterns:
    def normalize_section(section):
        if not section or pd.isna(section):
            return section
        section = str(section)
        # If section matches pattern like AI-A1, AI-B2, etc., keep as is
        if re.match(r'^[A-Z]{2,3}-[A-Z]\d+$', section):
            return section
        # Otherwise normalize other section formats
        if '-' in section and len(section.split('-')[1]) > 1:
            return section.split('-')[0] + '-' + section.split('-')[1][0]
        return section

    df['Section'] = df['Section'].apply(normalize_section)
    
    # 3. Convert times to proper datetime for sorting
    def get_start_time(time_str):
        if pd.isna(time_str):
            return pd.to_datetime('23:59', format='%H:%M')  # Put invalid times at end
        
        try:
            start_time = time_str.split('-')[0].strip()
            
            # Parse the time
            time_obj = pd.to_datetime(start_time, format='%H:%M')
            hour = time_obj.hour
            minute = time_obj.minute
            
            # Since classes run from 8:30 AM to 5:15 PM (17:15)
            # Any time before 8:30 should be interpreted as PM (add 12 hours)
            if hour < 8 or (hour == 8 and minute < 30):
                # This is PM time, convert to 24-hour format
                if hour != 12:  # Don't add 12 to 12 PM
                    hour += 12
                time_obj = pd.to_datetime(f'{hour:02d}:{minute:02d}', format='%H:%M')
            
            return time_obj
            
        except Exception as e:
            print(f"Error parsing time '{time_str}': {e}")
            return pd.to_datetime('23:59', format='%H:%M')  # Put invalid times at end
    
    # 4. Calculate class duration in minutes
    def calculate_duration(time_str):
        if pd.isna(time_str) or '-' not in str(time_str):
            return 0
        try:
            start_time, end_time = str(time_str).split('-')
            start_parts = start_time.strip().split(':')
            end_parts = end_time.strip().split(':')
            
            start_minutes = int(start_parts[0]) * 60 + int(start_parts[1])
            end_minutes = int(end_parts[0]) * 60 + int(end_parts[1])
            
            # Handle PM times (times before 8:30 are PM)
            start_hour = int(start_parts[0])
            end_hour = int(end_parts[0])
            
            if start_hour < 8 or (start_hour == 8 and int(start_parts[1]) < 30):
                start_minutes += 12 * 60
            if end_hour < 8 or (end_hour == 8 and int(end_parts[1]) < 30):
                end_minutes += 12 * 60
                
            return end_minutes - start_minutes
        except:
            return 0
    
    df['Duration'] = df['Class Time'].apply(calculate_duration)
    df['StartTime'] = df['Class Time'].apply(get_start_time)
    
    # 5. Remove duplicate classes, keeping the one with longest duration
    def deduplicate_classes(group):
        if len(group) == 1:
            return group
        # For duplicate classes, keep the one with maximum duration
        return group.loc[group['Duration'].idxmax()].to_frame().T
    
    # Group by Day, Course Name, Room, Section, Batch and keep longest duration
    df = df.groupby(['Day', 'Course Name', 'Room No', 'Section', 'Batch'], dropna=False).apply(deduplicate_classes).reset_index(drop=True)
    
    return df

timetable_df = preprocess_timetable()

@app.route('/')
def index():
    # Get unique days, batches, and sections for the dropdowns
    days = sorted(timetable_df['Day'].unique())
    batches = sorted(timetable_df['Batch'].dropna().unique())
    sections = sorted(timetable_df['Section'].dropna().unique())
    
    return render_template('index.html', days=days, batches=batches, sections=sections)

@app.route('/get_filtered_timetable', methods=['POST'])
def get_filtered_timetable():
    day = request.form.get('day')
    batch = request.form.get('batch')
    section = request.form.get('section')
    class_type = request.form.get('class_type', 'All')
    
    # Filter the timetable based on selections
    filtered_df = timetable_df.copy()
    
    if day and day != 'All':
        filtered_df = filtered_df[filtered_df['Day'] == day]
    
    if batch and batch != 'All':
        filtered_df = filtered_df[filtered_df['Batch'] == batch]
    
    if section and section != 'All':
        filtered_df = filtered_df[filtered_df['Section'] == section]
    
    if class_type and class_type != 'All':
        filtered_df = filtered_df[filtered_df['Type'] == class_type]
    
    # Sort by time
    filtered_df = filtered_df.sort_values('StartTime')
    
    # Convert to HTML table (drop Duration column as well)
    display_columns = ['Day', 'Course Name', 'Class Time', 'Room No', 'Section', 'Batch', 'Type']
    html_table = filtered_df[display_columns].to_html(
        classes='timetable-table', 
        index=False
    )
    
    return jsonify({
        'html': html_table,
        'count': len(filtered_df)
    })

@app.route('/get_sections', methods=['POST'])
def get_sections():
    batch = request.form.get('batch')
    
    if batch == 'All':
        sections = sorted(timetable_df['Section'].dropna().unique())
    else:
        sections = sorted(timetable_df[timetable_df['Batch'] == batch]['Section'].dropna().unique())
    
    return jsonify({'sections': sections})

if __name__ == '__main__':
    app.run(debug=True)