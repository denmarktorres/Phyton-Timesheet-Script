import smartsheet
import os

# Step 1: Set up API connection
API_TOKEN = "6QwQAHAdwi88DkTpWh47ixYtCMf8kRDCt0MlY"  # Replace with your Smartsheet API token
SHEET_ID = "8817580257005444"  # Replace with your Smartsheet sheet ID
COLUMNS_TO_SUM = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]  # Replace with columns to sum
HELPER_TYPE_COLUMN = "Helper Type"  # Replace with the Helper Type column name
EMAIL_COLUMN = "Email"
TIMESHEET_WEEK_COLUMN = "Timesheet Week"

# Initialize Smartsheet client
smartsheet_client = smartsheet.Smartsheet(API_TOKEN)

# Helper: Get all rows in the sheet
def get_all_rows(sheet_id):
    try:
        sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
        print(f"Successfully fetched sheet with ID: {sheet_id}")
        return sheet.rows, {col.title: col.id for col in sheet.columns}
    except Exception as e:
        print(f"Error fetching sheet: {e}")
        return [], {}

# Helper: Find existing Helper Rows
def find_helper_row(rows, email, timesheet_week, helper_type_col_id, email_col_id, week_col_id):
    for row in rows:
        helper_type = next((cell.value for cell in row.cells if cell.column_id == helper_type_col_id), None)
        row_email = next((cell.value for cell in row.cells if cell.column_id == email_col_id), None)
        row_week = next((cell.value for cell in row.cells if cell.column_id == week_col_id), None)

        if helper_type == "Helper" and row_email == email and row_week == timesheet_week:
            return row
    return None

# Helper: Calculate sums for user entries (excluding Helper rows)
def calculate_sums(rows, email, timesheet_week, email_col_id, week_col_id, sum_col_ids, helper_type_col_id):
    sums = {col_id: 0 for col_id in sum_col_ids}
    for row in rows:
        # Skip rows that are Helper type
        helper_type = next((cell.value for cell in row.cells if cell.column_id == helper_type_col_id), None)
        if helper_type == "Helper":
            continue
        
        row_email = next((cell.value for cell in row.cells if cell.column_id == email_col_id), None)
        row_week = next((cell.value for cell in row.cells if cell.column_id == week_col_id), None)

        if row_email == email and row_week == timesheet_week:
            for col_id in sum_col_ids:
                cell_value = next((cell.value for cell in row.cells if cell.column_id == col_id), 0)
                if isinstance(cell_value, (int, float)):
                    sums[col_id] += cell_value
    return sums

# Step 2: Add or update a helper row
def add_or_update_helper_row(sheet_id, email, timesheet_week, sums, column_map, existing_helper_row=None):
    if existing_helper_row:
        # If helper row exists, delete it first
        try:
            smartsheet_client.Sheets.delete_rows(sheet_id, [existing_helper_row.id])
            print(f"Existing helper row deleted for Email: {email}, Week: {timesheet_week}.")
        except Exception as e:
            print(f"Error deleting existing helper row for Email: {email}, Week: {timesheet_week}: {e}")

    # Create a new helper row to be added at the bottom
    new_row = smartsheet.models.Row()
    new_row.to_bottom = True  # Ensure the row is added at the bottom of the sheet

    # Add cells for the helper row
    new_row.cells.append({
        "column_id": column_map[EMAIL_COLUMN],
        "value": email
    })
    new_row.cells.append({
        "column_id": column_map[TIMESHEET_WEEK_COLUMN],
        "value": timesheet_week
    })
    new_row.cells.append({
        "column_id": column_map[HELPER_TYPE_COLUMN],
        "value": "Helper"
    })
    for col_name, col_id in column_map.items():
        if col_id in sums:
            new_row.cells.append({
                "column_id": col_id,
                "value": sums[col_id]
            })

    # Add the row to the sheet
    try:
        response = smartsheet_client.Sheets.add_rows(sheet_id, [new_row])
        print(f"Row successfully added to Smartsheet for Email: {email}, Week: {timesheet_week}.")
        print(f"Response: {response}")
    except Exception as e:
        print(f"Error adding row to Smartsheet for Email: {email}, Week: {timesheet_week}: {e}")

# Step 3: Main logic
def main():
    # Fetch rows and column map
    rows, column_map = get_all_rows(SHEET_ID)
    
    if not rows or not column_map:
        print("Unable to fetch rows or column map.")
        return

    # Debug: Print column map to verify columns
    print(f"Column Map: {column_map}")

    helper_type_col_id = column_map.get(HELPER_TYPE_COLUMN)
    email_col_id = column_map.get(EMAIL_COLUMN)
    week_col_id = column_map.get(TIMESHEET_WEEK_COLUMN)
    sum_col_ids = [column_map.get(col) for col in COLUMNS_TO_SUM]

    # Check if any required columns are missing
    if None in [helper_type_col_id, email_col_id, week_col_id] + sum_col_ids:
        print("Error: One or more required columns are missing.")
        return

    # Loop through each row and check if helper row is needed
    processed_combinations = set()  # Keep track of email-week combinations we've processed
    for row in rows:
        email = next((cell.value for cell in row.cells if cell.column_id == email_col_id), None)
        timesheet_week = next((cell.value for cell in row.cells if cell.column_id == week_col_id), None)

        if email and timesheet_week:
            print(f"Processing row for Email: {email}, Timesheet Week: {timesheet_week}")

            # Only proceed if we haven't already processed this email-week combination
            if (email, timesheet_week) not in processed_combinations:
                # Check if a helper row already exists
                existing_helper_row = find_helper_row(rows, email, timesheet_week, helper_type_col_id, email_col_id, week_col_id)

                # Calculate sums for this user entry, excluding Helper rows
                sums = calculate_sums(rows, email, timesheet_week, email_col_id, week_col_id, sum_col_ids, helper_type_col_id)

                # Add or update the helper row
                add_or_update_helper_row(SHEET_ID, email, timesheet_week, sums, column_map, existing_helper_row)
                
                # Mark this email-week combination as processed
                processed_combinations.add((email, timesheet_week))
            else:
                print(f"Skipping processing for Email: {email}, Timesheet Week: {timesheet_week} as already processed")
        else:
            print(f"Skipping row as Email or Timesheet Week is missing for row: {row}")

if __name__ == "__main__":
    main()
