import os
import re
import openpyxl
import numpy as np
import pandas as pd
from typing import Dict, List, Set, Any, Optional

def clean_excel_cell(value: Any) -> Optional[Any]:
    """
    Cleans a cell value by stripping whitespace from strings and converting
    empty strings or whitespace-only strings to None.

    Args:
        value (Any): The raw value from an Excel cell.

    Returns:
        Optional[Any]: The cleaned value, or None if the original was None,
                       an empty string, or a whitespace-only string.
    """
    if value is None:
        return None
    if isinstance(value, str):
        cleaned_str = value.strip()
        return cleaned_str if cleaned_str != '' else None
    return value

def extract_sheet_data(workbook: openpyxl.workbook.workbook.Workbook, sheet_name: str) -> pd.DataFrame:
    """
    Processes a single sheet from an OpenPyXL workbook to extract tabular data.
    
    This function handles merged cells by propagating their values and dynamically
    identifies the header row and the last meaningful data column.
    It also ensures that empty or whitespace-only cells are treated as None/NaN.

    Args:
        workbook (openpyxl.workbook.workbook.Workbook): The OpenPyXL workbook object.
        sheet_name (str): The name of the sheet to process within the workbook.

    Returns:
        pd.DataFrame: A pandas DataFrame containing the cleaned tabular data,
                      or an empty DataFrame if the 'SLOT' header is not found
                      or no meaningful data columns are present.
    """
    sheet = workbook[sheet_name]
    merged_ranges = sheet.merged_cells.ranges
    
    # Create a new in-memory workbook and sheet to copy data into.
    # This is done to "unmerge" cells by copying their values to all cells
    # within their original merged range, making data extraction easier.
    new_workbook = openpyxl.Workbook()
    new_sheet = new_workbook.active

    # Copy all cell data to the new sheet, cleaning values during transfer.
    for row in sheet.rows:
        for cell in row:
            cleaned_value = clean_excel_cell(cell.value)
            new_sheet.cell(row=cell.row, column=cell.column, value=cleaned_value)
    
    # Propagate values for merged cells in the new sheet.
    for merged_range in merged_ranges:
        # Get the cleaned value from the top-left cell of the original merged range.
        value_to_propagate = clean_excel_cell(sheet.cell(merged_range.min_row, merged_range.min_col).value)
        
        # Fill all cells within the merged range in the new_sheet with this value.
        for row_index in range(merged_range.min_row, merged_range.max_row + 1):
            for col_index in range(merged_range.min_col, merged_range.max_col + 1):
                new_sheet.cell(row=row_index, column=col_index, value=value_to_propagate)
    
    # Convert the processed sheet to a pandas DataFrame.
    # Empty cells (None) will be converted to NaN by pandas.
    data_frame = pd.DataFrame(new_sheet.values)

    # --- Dynamic Header Row Detection ---
    # Attempt to find the header row by looking for 'SLOT' in the second column (index 1).
    try:
        # Convert the second column to string type to ensure robust comparison,
        # as it might contain mixed data types.
        header_index = data_frame[data_frame.iloc[:, 1].astype(str) == 'SLOT'].index[0]
    except IndexError:
        print(f"Warning: 'SLOT' not found in the second column of sheet '{sheet_name}'. Returning empty DataFrame.")
        return pd.DataFrame()
        
    header_row_values = data_frame.iloc[header_index]
    
    # --- Dynamically Trim Columns based on Header Row ---
    # Find the first column in the header row that is NaN/None.
    # This indicates the boundary where meaningful data columns end.
    first_unwanted_col_index = -1
    for i, col_value in enumerate(header_row_values):
        if pd.isna(col_value):
            first_unwanted_col_index = i
            break
            
    # If an unwanted column was found, slice the DataFrame to keep only meaningful columns.
    if first_unwanted_col_index != -1:
        data_frame = data_frame.iloc[:, :first_unwanted_col_index]
    
    # Set the DataFrame columns using the identified header row values.
    # Slice header_row_values to match the actual number of columns in the trimmed DataFrame.
    data_frame.columns = header_row_values.iloc[:data_frame.shape[1]]
    
    # Remove the header row and all rows above it, then reset the index.
    data_frame = data_frame.iloc[header_index + 1:].reset_index(drop=True)
    
    # --- Final Cleanup - Drop entirely empty rows ---
    # Replace any remaining empty strings with NaN (a safeguard, as clean_excel_cell should handle most).
    data_frame.replace('', np.nan, inplace=True)
    # Drop rows where all values are NaN.
    data_frame.dropna(how='all', inplace=True)
    
    return data_frame

def extract_subject_details(subject_string: str, subject_abbreviations: Set[str]) -> Optional[Dict[str, Any]]:
    """
    Parses a raw subject string (e.g., "SUB1 5A3/B3") into its components.

    It extracts the subject code, semester, and a list of division-batch objects.
    This function assumes the format: "SUBJECT_CODE [SEMESTER][DIVISIONS][BATCH]".
    It filters out 'TUT' (tutorial) entries and validates the subject_code
    against a set of known subject abbreviations.

    Args:
        subject_string (str): The raw subject string from the processed DataFrame cell.
        subject_abbreviations (Set[str]): A set of strings containing all valid
                                          subject codes to consider.

    Returns:
        Optional[Dict[str, Any]]: A dictionary containing 'subject_code', 'semester', and
                                  'division_batches' if parsing is successful. Returns None
                                  if the input is invalid, a tutorial, does not match the
                                  expected format, or if the subject_code is not in the
                                  provided `subject_abbreviations` set.
                                  Example output:
                                  {
                                      'subject_code': 'OT',
                                      'semester': 5,
                                      'division_batches': [{'division': 'A', 'batch': '3'}, {'division': 'B', 'batch': '3'}]
                                  }
    """
    # Initial validation: check if input is a valid non-empty string and not a tutorial.
    if not isinstance(subject_string, str) or not subject_string.strip() or 'TUT' in subject_string.upper():
        return None

    # Split the string into parts based on whitespace.
    # Expected format: ['SUBJECT_CODE', 'CLASS_INFO']
    # maxsplit=1 ensures only the first space splits, handling subject codes with spaces if any.
    parts = subject_string.strip().split(maxsplit=1)

    # Ensure there are at least two parts (subject code and class info).
    if len(parts) < 2:
        return None

    subject_code = parts[0]

    # Validate subject_code against known abbreviations.
    if subject_code not in subject_abbreviations:
        return None

    class_info = parts[1] # This part contains semester, divisions, and batch (e.g., "5A3/B3")

    semester: Optional[int] = None
    division_batches: List[Dict[str, Optional[str]]] = []

    # --- Extract Semester ---
    # Use regex to find the leading digit(s) for the semester.
    semester_match = re.match(r'^(\d+)', class_info)
    if semester_match:
        semester = int(semester_match.group(1))
        # Remove the semester prefix from class_info for subsequent parsing.
        class_info_without_semester = class_info[semester_match.end():]
    else:
        # If no semester digit is found at the beginning, the format is unexpected.
        return None

    # --- Process Divisions and Batch ---
    # Handle the special 'ALL' division case.
    if 'ALL' in class_info_without_semester.upper():
        division_batches = [{'division': 'ALL', 'batch': None}]
    else:
        # Split the remaining class_info by '/' to get individual division segments
        # (e.g., "A3", "B3", "A", "B*").
        division_segments = class_info_without_semester.split('/')

        for segment in division_segments:
            # Extract only alphabetic characters for the division letter, preserving case.
            division_letter = ''.join(char for char in segment if char.isalpha()).strip()

            batch_value: Optional[str] = None
            # Regex looks for one or more digits followed by an optional asterisk at the end.
            batch_match = re.search(r'(\d+\*?)$', segment)
            if batch_match:
                batch_value = batch_match.group(1)

            if division_letter: # Only add if a valid division letter was found.
                division_batches.append({'division': division_letter, 'batch': batch_value})

    # Ensure division_batches are unique (based on division and batch combined)
    # and sorted for consistent output, unless 'ALL' is the only entry.
    if not division_batches or division_batches[0].get('division') != 'ALL':
        unique_division_batches: List[Dict[str, Optional[str]]] = []
        seen_tuples: Set[tuple] = set()
        for item in division_batches:
            item_tuple = (item['division'], item['batch'])
            if item_tuple not in seen_tuples:
                unique_division_batches.append(item)
                seen_tuples.add(item_tuple)

        # Sort by division letter, then by batch (None values sorted last).
        division_batches = sorted(unique_division_batches, key=lambda x: (x['division'], x['batch'] if x['batch'] is not None else ''))

    return {
        'subject_code': subject_code,
        'semester': semester,
        'division_batches': division_batches,
    }

def build_faculty_schedules(processed_data_frame: pd.DataFrame, known_faculty_abbreviations: Set[str], subject_abbreviations: Set[str]) -> Dict[str, Dict[str, List[Dict[str, Any]]]]:
    """
    Creates a dictionary mapping each faculty member to their weekly schedule.

    This optimized version now includes the parsed subject information directly
    within each session entry, avoiding redundant parsing later.

    Args:
        processed_data_frame (pd.DataFrame): The DataFrame processed by `extract_sheet_data`.
        known_faculty_abbreviations (Set[str]): A set of valid faculty abbreviations.
        subject_abbreviations (Set[str]): A set of valid subject abbreviations to use
                                     when parsing subject strings.

    Returns:
        Dict[str, Dict[str, List[Dict[str, Any]]]]: A dictionary where keys are faculty names
                                                    and values are dictionaries containing their
                                                    schedule for each day of the week.
    """
    
    DAY_MAPPING: Dict[str, str] = {
        'MON': 'Monday', 'MONDAY': 'Monday',
        'TUES': 'Tuesday', 'TUESDAY': 'Tuesday',
        'WED': 'Wednesday', 'WEDNESDAY': 'Wednesday',
        'THUR': 'Thursday', 'THURSDAY': 'Thursday',
        'FRI': 'Friday', 'FRIDAY': 'Friday',
        'SAT': 'Saturday', 'SATURDAY': 'Saturday'
    }

    faculty_master: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}

    # Identify actual faculty names from the DataFrame columns, excluding the first two
    # (assumed to be 'Day' and 'SLOT') and filtering by known abbreviations.
    faculty_names = [
        col for col in processed_data_frame.columns[2:]
        if col in known_faculty_abbreviations
    ]

    # Initialize an empty schedule for each identified faculty member for all days.
    for faculty in faculty_names:
        faculty_master[faculty] = {
            'Monday': [], 'Tuesday': [], 'Wednesday': [],
            'Thursday': [], 'Friday': [], 'Saturday': []
        }

    # Process schedules for each identified faculty member.
    for faculty in faculty_names:
        faculty_column = processed_data_frame[faculty] # Get the series for the current faculty's schedule

        row_index = 0 # Correctly initialized loop variable
        while row_index < len(faculty_column):
            raw_day_value = processed_data_frame.iloc[row_index, 0]
            raw_time_slot_start = processed_data_frame.iloc[row_index, 1]
            raw_subject_string = faculty_column.iloc[row_index]

            # Process day value.
            day_string = str(raw_day_value).strip().upper()
            day = DAY_MAPPING.get(day_string)

            # Check if cell contains valid data and a valid day was mapped.
            if pd.notna(raw_subject_string) and raw_subject_string != '' and day is not None:
                # Convert time_slot_start to int, with error handling.
                try:
                    time_slot_start = int(raw_time_slot_start)
                except (ValueError, TypeError):
                    print(f"Warning: Time slot '{raw_time_slot_start}' for subject '{raw_subject_string}' "
                          f"for faculty '{faculty}' on {day} is not a valid integer. Skipping entry.")
                    row_index += 1
                    continue

                current_subject_value = raw_subject_string
                block_start_row_index = row_index

                # Find the end of the contiguous block for the current subject.
                block_end_row_index = row_index # Correct variable name
                while (block_end_row_index + 1 < len(faculty_column) and
                       faculty_column.iloc[block_end_row_index + 1] == current_subject_value and # Correct variable name
                       str(processed_data_frame.iloc[block_end_row_index + 1, 0]).strip().upper() == day_string): # Correct variable name
                    block_end_row_index += 1 # Correct variable name

                # Calculate the length of the contiguous block in terms of rows/slots.
                block_length = (block_end_row_index - block_start_row_index) + 1

                # Determine activity type based on fixed slot allocation rules.
                activity_type: str
                if block_length == 2:
                    activity_type = 'Lab'
                elif block_length == 1:
                    activity_type = 'Lecture'
                else:
                    activity_type = 'Unknown'
                    print(f"Warning: Unexpected block length ({block_length}) for subject '{current_subject_value}' "
                          f"for faculty '{faculty}' on {day} at slot {time_slot_start}. Classified as '{activity_type}'.")

                # Store the raw subject string and its parsed information.
                schedule_entry = {
                    'subject_string': current_subject_value,
                    'type': activity_type,
                    'time_slot': time_slot_start,
                    # Pass subject_abbreviations to extract_subject_details
                    'parsed_subject_info': extract_subject_details(current_subject_value, subject_abbreviations)
                }

                faculty_master[faculty][day].append(schedule_entry)

                # Advance row_index to the end of the processed block to avoid re-processing.
                row_index = block_end_row_index + 1
            else:
                # If the cell is empty, invalid, or day mapping failed, move to the next row.
                row_index += 1

    return faculty_master

def standardize_time_slots(data_frame: pd.DataFrame) -> pd.DataFrame:
    """
    Formats the 'Time_Slot' column for 'Lab' entries in a DataFrame.
    
    For 'Lab' entries, it converts the single starting time slot (e.g., 3)
    into a two-slot range string (e.g., "3-4"). Lecture time slots remain as is.
    The DataFrame is then sorted.

    Args:
        data_frame (pd.DataFrame): The DataFrame containing schedule entries.
                           Expected columns include 'Type' and 'Time_Slot'.

    Returns:
        pd.DataFrame: The DataFrame with 'Time_Slot' formatted for labs,
                      and sorted by 'Day', 'Time_Slot', and 'Batch'.
    """
    df_copy = data_frame.copy() 

    # Explicitly cast 'Time_Slot' to object dtype to allow mixed types (int and str)
    # before applying string formatting. This resolves the FutureWarning.
    df_copy['Time_Slot'] = df_copy['Time_Slot'].astype(object)

    lab_mask = df_copy['Type'] == 'Lab'
    df_copy.loc[lab_mask, 'Time_Slot'] = df_copy.loc[lab_mask, 'Time_Slot'].apply(
        lambda x: f"{int(x)}-{int(x)+1}" if pd.notna(x) else x
    )
    
    def get_time_slot_sort_value(slot: Any) -> Any:
        """Helper function to extract a sortable value from time slot (int or 'start-end' string)."""
        if isinstance(slot, str) and '-' in slot:
            try:
                # For ranges like "3-4", sort by the starting number.
                return int(slot.split('-')[0])
            except ValueError:
                # Handle cases where string is not a valid range (e.g., "invalid-slot"), sort at end.
                return float('inf') 
        elif pd.isna(slot):
            # Put NaN values at the end of the sort order.
            return float('inf')
        try:
            # For single integer slots.
            return int(slot)
        except (ValueError, TypeError):
            # Handle non-numeric or unconvertible slots, sort at end.
            return float('inf')

    # Sort the DataFrame.
    # The 'Time_Slot' column needs a custom key because it will contain mixed types (int and str).
    # 'Batch' might contain None, so it's converted to string for consistent sorting.
    return df_copy.sort_values(
        by=['Day', 'Time_Slot', 'Batch'],
        key=lambda col: col.apply(get_time_slot_sort_value) if col.name == 'Time_Slot' else col.astype(str),
        ignore_index=True
    )

def generate_class_schedules(faculty_schedules: Dict[str, Dict[str, List[Dict[str, Any]]]]) -> Dict[str, pd.DataFrame]:
    """
    Creates individual timetable DataFrames for each class division.

    It processes the consolidated faculty schedules, extracts subject details,
    and assigns sessions to the relevant divisions (including 'ALL' divisions).
    This optimized version uses pre-parsed subject information stored within
    the faculty schedules.

    Args:
        faculty_schedules (Dict[str, Dict[str, List[Dict[str, Any]]]]): A dictionary of
                                                                         consolidated faculty schedules,
                                                                         as returned by `process_all_timetables`.
                                                                         Each session entry is expected to
                                                                         contain a 'parsed_subject_info' key.

    Returns:
        Dict[str, pd.DataFrame]: A dictionary where keys are division identifiers
                                 (e.g., "5A", "6B") and values are pandas DataFrames
                                 representing the timetable for that division.
                                 The DataFrames are formatted and sorted by
                                 `standardize_time_slots`.
    """
    division_tables: Dict[str, List[Dict[str, Any]]] = {}
    semester_divisions: Dict[int, Set[str]] = {} # To store all unique divisions for each semester (for 'ALL' cases)

    # --- First Pass: Collect all unique divisions per semester ---
    # This pass is crucial for correctly expanding 'ALL' entries later.
    for faculty_name, faculty_schedule_by_day in faculty_schedules.items():
        for day, sessions in faculty_schedule_by_day.items():
            for session_entry in sessions:
                # Retrieve pre-parsed subject info, which should be available
                # from the `build_faculty_schedules` function.
                parsed_subject_info = session_entry.get('parsed_subject_info')

                # Only proceed if parsing was successful and semester information is present.
                if parsed_subject_info and parsed_subject_info['semester'] is not None:
                    semester = parsed_subject_info['semester']
                    
                    # Initialize the set for the current semester if it doesn't exist.
                    if semester not in semester_divisions:
                        semester_divisions[semester] = set()
                    
                    # Iterate through the division_batches to collect individual divisions.
                    # We only add specific divisions, not the 'ALL' placeholder itself.
                    for db_entry in parsed_subject_info.get('division_batches', []):
                        division_letter = db_entry.get('division')
                        if division_letter and division_letter != 'ALL':
                            semester_divisions[semester].add(division_letter)
    
    # --- Second Pass: Populate division tables ---
    # Iterate through each faculty's consolidated schedule to assign sessions to divisions.
    for faculty_name, faculty_schedule_by_day in faculty_schedules.items():
        for day, sessions in faculty_schedule_by_day.items():
            for session_entry in sessions:
                # Retrieve pre-parsed subject info for the current session.
                parsed_subject_info = session_entry.get('parsed_subject_info')
                
                # Only process if subject parsing was successful and semester is valid.
                if parsed_subject_info and parsed_subject_info['semester'] is not None:
                    semester = parsed_subject_info['semester']
                    target_division_batch_entries: List[Dict[str, Optional[str]]] = []
                    
                    # Determine the target divisions/batches for the current session.
                    session_division_batches = parsed_subject_info.get('division_batches', [])

                    # Check if the first entry indicates 'ALL' divisions.
                    if session_division_batches and session_division_batches[0].get('division') == 'ALL':
                        # If 'ALL' is specified, expand to all unique divisions known for this semester.
                        if semester in semester_divisions:
                            # Sort divisions for consistent output order.
                            for div in sorted(list(semester_divisions[semester])):
                                # For 'ALL' subjects, the batch is typically not specific to a sub-division.
                                target_division_batch_entries.append({'division': div, 'batch': None})
                    else:
                        # Otherwise, use the specific division-batch combinations parsed from the subject string.
                        target_division_batch_entries = session_division_batches
                    
                    # Add the session entry to each relevant division's timetable.
                    for db_entry in target_division_batch_entries:
                        division = db_entry.get('division')
                        batch = db_entry.get('batch') # Get the batch specific to this division entry.
                        
                        if division: # Ensure division is not None or empty
                            # Construct the unique key for the division table (e.g., "5A", "6B").
                            division_key = f"{semester}{division}"
                            
                            # Initialize the list for this division key if it doesn't exist.
                            if division_key not in division_tables:
                                division_tables[division_key] = []
                            
                            # Create the entry for the division's timetable.
                            entry = {
                                'Subject': parsed_subject_info.get('subject_code'),
                                'Type': session_entry.get('type'), # Use type directly from faculty_schedules.
                                'Batch': batch if batch is not None else '-', # Use specific batch, or '-' if None.
                                'Day': day,
                                'Time_Slot': session_entry.get('time_slot'), # Use the single time_slot.
                                'Faculty': faculty_name
                            }
                            
                            division_tables[division_key].append(entry)
    
    # --- Convert lists of entries to DataFrames and apply final formatting ---
    final_division_dataframes: Dict[str, pd.DataFrame] = {}
    for division_key, entries_list in division_tables.items():
        if entries_list: # Only create DataFrame if there are entries.
            data_frame = pd.DataFrame(entries_list)
            # Apply standardize_time_slots to format time slots and sort the DataFrame.
            final_division_dataframes[division_key] = standardize_time_slots(data_frame)
        else:
            # If no entries for a division, return an empty DataFrame with expected columns.
            final_division_dataframes[division_key] = pd.DataFrame(columns=['Subject', 'Type', 'Batch', 'Day', 'Time_Slot', 'Faculty'])
    
    return final_division_dataframes

def process_all_timetables(matrix_file_path: str, faculties: Set[str], subjects: Set[str]) -> Dict[str, Dict[str, List[Dict[str, Any]]]]:
    """
    Loads an Excel workbook, processes each sheet, and consolidates
    faculty schedules from all sheets into a single dictionary.

    Args:
        matrix_file_path (str): The path to the Excel file containing timetable data.
        faculties (Set[str]): A set of strings containing all valid
                               faculty abbreviations to consider as columns.
        subjects (Set[str]): A set of strings containing all valid
                             subject abbreviations for parsing.

    Returns:
        Dict[str, Dict[str, List[Dict[str, Any]]]]: A consolidated dictionary of all faculty schedules
                                                    across all sheets, with entries sorted by time slot
                                                    for each day. Returns an empty dictionary if the
                                                    file is not found or an error occurs during loading.
    """
    try:
        workbook = openpyxl.load_workbook(matrix_file_path)
    except FileNotFoundError:
        print(f"Error: The file '{matrix_file_path}' was not found.")
        return {}
    except Exception as e:
        print(f"Error loading workbook '{matrix_file_path}': {e}")
        return {}

    all_faculty_schedules: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}

    # Iterate through each sheet in the workbook.
    for sheet_name in workbook.sheetnames:
        # print(f"Processing sheet: {sheet_name}")
        
        # Process the current sheet into a cleaned pandas DataFrame.
        # This function (extract_sheet_data) is assumed to be defined elsewhere.
        processed_data_frame = extract_sheet_data(workbook, sheet_name)
        
        # Only proceed if the processed DataFrame contains data.
        if not processed_data_frame.empty:
            # Create a schedule dictionary for the current sheet.
            # This function (build_faculty_schedules) is assumed to be defined elsewhere.
            # It now requires both faculty and subject abbreviations.
            sheet_schedules = build_faculty_schedules(
                processed_data_frame, faculties, subjects
            )
            
            # Merge schedules from the current sheet into the overall consolidated schedules.
            for faculty_name, schedule_by_day in sheet_schedules.items():
                if faculty_name not in all_faculty_schedules:
                    # If the faculty is encountered for the first time, add their entire schedule.
                    all_faculty_schedules[faculty_name] = schedule_by_day
                else:
                    # If the faculty already exists, extend their daily schedules with new sessions.
                    for day_name in schedule_by_day:
                        all_faculty_schedules[faculty_name][day_name].extend(schedule_by_day[day_name])
        else:
            print(f"Skipping sheet '{sheet_name}' due to empty or invalid data.")

    # --- Post-processing: Sort schedule entries by time_slot for each faculty and day ---
    # This ensures a consistent and chronologically ordered output for each faculty's daily schedule.
    for faculty_name, schedule_by_day in all_faculty_schedules.items():
        for day_name, entries in schedule_by_day.items():
            # Only attempt to sort if there are entries and the 'time_slot' key is present.
            if entries and 'time_slot' in entries[0]:
                all_faculty_schedules[faculty_name][day_name] = sorted(entries, key=lambda x: x['time_slot'])
            # If 'time_slot' is not consistently present (e.g., due to 'Unknown' types),
            # or if the list is empty, no sorting is applied to avoid errors.
            
    return all_faculty_schedules

def get_division_course_catalog(division_tables: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    """
    Creates condensed timetable DataFrames for each class division.

    This function takes the detailed division timetables and removes
    time-related information ('Time_Slot', 'Day') to provide a unique
    list of subjects, their types, batches, and associated faculty for
    each division.

    Args:
        division_tables (Dict[str, pd.DataFrame]): A dictionary where keys are division identifiers
                                                   (e.g., "5A", "6B") and values are pandas DataFrames
                                                   representing the detailed timetable for that division.
                                                   These DataFrames are expected to have columns like
                                                   'Subject', 'Type', 'Batch', 'Day', 'Time_Slot', 'Faculty'.

    Returns:
        Dict[str, pd.DataFrame]: A dictionary where keys are division identifiers
                                 and values are pandas DataFrames representing the
                                 condensed timetable for that division. These DataFrames
                                 will contain 'Subject', 'Type', 'Batch', and 'Faculty'
                                 columns, with duplicate entries removed and sorted.
    """
    condensed_tables: Dict[str, pd.DataFrame] = {}
    
    # Iterate through each division's detailed timetable DataFrame.
    for division_key, data_frame in division_tables.items():
        # --- Step 1: Drop time-related columns ---
        # Create a new DataFrame by dropping the 'Time_Slot' and 'Day' columns.
        # This removes the temporal information, making the table "condensed".
        # `axis=1` specifies that columns should be dropped.
        # `errors='ignore'` prevents an error if 'Time_Slot' or 'Day' columns are already missing.
        condensed_df = data_frame.drop(['Time_Slot', 'Day'], axis=1, errors='ignore')
        
        # --- Step 2: Remove duplicate entries ---
        # After dropping 'Time_Slot' and 'Day', multiple rows might become identical
        # (e.g., if a subject is taught on different days or at different times).
        # `drop_duplicates()` ensures that each unique combination of 'Subject',
        # 'Type', 'Batch', and 'Faculty' appears only once.
        condensed_df = condensed_df.drop_duplicates()
        
        # --- Step 3: Sort for better readability ---
        # Sort the condensed DataFrame by 'Subject' first, then by 'Batch'.
        # This organizes the output logically for easier review.
        # `reset_index(drop=True)` creates a new default integer index after sorting.
        # The `key` argument is used for 'Batch' to ensure consistent sorting
        # even if 'Batch' contains mixed types (e.g., '1', '2', '-', None).
        condensed_df = condensed_df.sort_values(
            by=['Subject', 'Batch'],
            key=lambda col: col.astype(str) if col.name == 'Batch' else col,
            ignore_index=True
        )
        
        # --- Step 4: Store the condensed DataFrame ---
        # Add the processed (condensed and unique) DataFrame to the result dictionary
        # using the original division key.
        condensed_tables[division_key] = condensed_df
    
    return condensed_tables

def build_hierarchical_schedule(condensed_division_tables: Dict[str, pd.DataFrame], department: str, college: str = "LDRP-ITR") -> Dict[str, Any]:
    """
    Creates a final hierarchical dictionary consolidating all timetable information
    grouped by college, department, semester, division, and subject.

    This function processes the condensed timetable DataFrames to extract
    designated faculty for lectures and labs (per batch).

    Args:
        condensed_division_tables (Dict[str, pd.DataFrame]): A dictionary where keys are
                                                             division identifiers (e.g., "5A", "6B")
                                                             and values are pandas DataFrames
                                                             representing the condensed timetable
                                                             for that division. These DataFrames
                                                             are expected to have 'Subject', 'Type',
                                                             'Batch', and 'Faculty' columns.
        department (str): The name of the department (e.g., "Computer Engineering").
        college (str, optional): The name of the college. Defaults to "LDRP-ITR".

    Returns:
        Dict[str, Any]: A nested dictionary containing the organized timetable data.
                        Example structure:
                        {
                            'LDRP-ITR': {
                                'Computer Engineering': {
                                    '5': {
                                        'A': {
                                            'AJP': {
                                                'lectures': {'designated_faculty': 'FacultyX'},
                                                'labs': {'1': {'designated_faculty': 'FacultyY'}}
                                            },
                                            'CN': { ... }
                                        },
                                        'B': { ... }
                                    }
                                }
                            }
                        }
    """
    # Initialize the main dictionary structure with college and department.
    final_consolidated_data: Dict[str, Any] = {
        college: {
            department: {}
        }
    }
    
    # Iterate through each condensed division timetable.
    for division_key, data_frame in condensed_division_tables.items():
        # Extract semester and division from the division_key (e.g., "5A" -> semester "5", division "A").
        # Assuming division_key format is always <semester_digit><division_letter(s)>.
        semester_str = division_key[0] # First character is the semester as a string.
        division_str = division_key[1:] # Rest of the string is the division.
        
        # Initialize semester entry if it doesn't exist within the department.
        if semester_str not in final_consolidated_data[college][department]:
            final_consolidated_data[college][department][semester_str] = {}
            
        # Initialize division entry if it doesn't exist within the semester.
        if division_str not in final_consolidated_data[college][department][semester_str]:
            final_consolidated_data[college][department][semester_str][division_str] = {}
            
        # Process each unique subject within the current division's DataFrame.
        # Using groupby('Subject') is efficient for processing subjects, as it groups
        # all rows pertaining to a single subject together.
        for subject_code, subject_data_group in data_frame.groupby('Subject'):
            subject_details: Dict[str, Any] = {
                'lectures': {},
                'labs': {}
            }
            
            # Process lectures for the current subject.
            lectures_df = subject_data_group[subject_data_group['Type'] == 'Lecture']
            if not lectures_df.empty:
                # Assuming one designated faculty for lectures per subject per division.
                # .iloc[0] picks the first faculty if multiple are listed (due to prior sorting).
                subject_details['lectures'] = {
                    'designated_faculty': lectures_df['Faculty'].iloc[0]
                }
            
            # Process labs for the current subject.
            labs_df = subject_data_group[subject_data_group['Type'] == 'Lab']
            if not labs_df.empty:
                # For labs, associate each batch with its designated faculty.
                # The 'Batch' column from the condensed DataFrame is used as the key.
                subject_details['labs'] = {
                    str(batch): {'designated_faculty': faculty} # Ensure batch is string for dictionary key.
                    for batch, faculty in zip(labs_df['Batch'], labs_df['Faculty'])
                }
            
            # Assign the processed subject data to the final hierarchical dictionary structure.
            final_consolidated_data[college][department][semester_str][division_str][subject_code] = subject_details
            
    return final_consolidated_data

def run_matrix_pipeline(matrix_file_path: str, faculty_abbreviations: Set[str], subject_abbreviations: Set[str], department: str, college: str = "LDRP-ITR") -> Dict[str, Any]:
    """
    Orchestrates the entire timetable processing pipeline to generate a final
    hierarchical dictionary of consolidated timetable information.

    This function calls a sequence of sub-functions to:
    1. Load and process Excel data into faculty-wise schedules.
    2. Transform faculty schedules into division-wise detailed timetables.
    3. Condense division timetables by removing time-slot and day information.
    4. Structure the condensed data into a final hierarchical dictionary.

    Args:
        matrix_file_path (str): The path to the Excel file containing timetable data.
        faculty_abbreviations (Set[str]): A set of valid faculty abbreviations.
        subject_abbreviations (Set[str]): A set of valid subject abbreviations.
        department (str): The name of the department (e.g., "Computer Engineering").
        college (str, optional): The name of the college. Defaults to "LDRP-ITR".

    Returns:
        Dict[str, Any]: A nested dictionary containing the organized timetable data
                        grouped by college, department, semester, division, and subject.
                        Returns an empty dictionary if any step in the pipeline fails
                        to produce valid data.
    """
    # print("Step 1: Generating full faculty schedules...")
    # Generate the full consolidated faculty schedules from the Excel file.
    # This includes initial data cleaning and subject parsing.
    all_faculty_schedules = process_all_timetables(
        matrix_file_path,
        faculties=faculty_abbreviations,
        subjects=subject_abbreviations
    )
    
    if not all_faculty_schedules:
        # print("Error: No faculty schedules generated. Aborting.")
        return {}

    # print("Step 2: Creating division-specific timetables...")
    # Create detailed timetable DataFrames for each class division.
    # This expands 'ALL' divisions and formats lab time slots.
    division_tables = generate_class_schedules(all_faculty_schedules)
    
    if not division_tables:
        # print("Error: No division timetables created. Aborting.")
        return {}

    # print("Step 3: Condensing division timetables...")
    # Create condensed tables for each division by removing time and day information,
    # and dropping duplicate subject entries.
    condensed_division_tables = get_division_course_catalog(division_tables)
    
    if not condensed_division_tables:
        print("Error: No condensed division timetables created. Aborting.")
        return {}

    # print("Step 4: Creating final hierarchical dictionary...")
    # Create the final hierarchical dictionary structure from the condensed tables.
    final_dict = build_hierarchical_schedule(
        condensed_division_tables,
        department=department,
        college=college
    )
    
    # print("Pipeline complete.")
    return final_dict
