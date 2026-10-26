# FT3 Excel workbook to per-test CSV file Conversion Script
#SL 10/22/2026 

import csv
import os
from datetime import datetime
import xlrd  # for reading Excel workbook

# ------------------- SETUP DIRECTORIES ------------------- 
script_dir = os.path.dirname(os.path.abspath(__file__))  # Get current directory
os.chdir(script_dir)  # Use the script folder as the current working directory.

# Keep all generated CSV files together in the "Results" folder"
output_results_dir = os.path.join(script_dir, "Results")  # Build path to new "Results" folder.
os.makedirs(output_results_dir, exist_ok=True)  # Create the Results folder if it is not already present.

print("Working directory:", script_dir)
print("Output folder:", output_results_dir)

#------------------- Creating Headers in the CSV ------------------------
# These CSV Header names define both the order and spelling of columns in every output CSV.
csv_headers = [  
    "Time and Date",  # Add 'Time and Date' as a CSV column name.
    "Tester",
    "Serial Number",
    "Overall Verdict",
    "FPGA Version",
    "Zigbee Version",
    "Aux Power Result", 
    "Grid Detect Result", 
    "MPPT Result", 
]  
# Generate the repeated [MPPT x]/[PV x] column names rather than typing all 48 fields
for mppt_number in range(1, 5):  # Repeat for MPPT numbers 1 through 4.
    for pv_number in range(1, 5):  # Repeat for each PV or condition number
        csv_headers.extend(  # Extend CSV header list with the MPPT/PV headers 
            [
                f"MPPT {mppt_number} PV {pv_number} Result",  # i.e. MPPT (1-4) PV (1-4) Result
                f"MPPT {mppt_number} PV {pv_number} Voltage",  # i.e. MPPT (1-4) PV (1-4) Voltage
                f"MPPT {mppt_number} PV {pv_number} Current",  # i.e. MPPT (1-4) PV (1-4) Current
            ]
        )
csv_headers.extend(  # Extend CSV header list with the remaining test result headers
    [  
        "AC Power",  
        "Maximum Power",  
        "GFDI Result",  
        "Zigbee Result",  
        "ABVF Result",  
        "ABVF 1",  
        "ABVF 2",  
        "ABVF 3", 
        "ABVF 4",  
    ]  
)

# ------------------- Creating new CSV Files ------------------------
def create_result_csv_files():
    
    date_formats = ( # Accepts dates in multiple formats, including the following:
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M",
        "%y/%m/%d %H:%M:%S",
        "%Y/%m/%d",
    )

    workbook_path = os.path.join(script_dir, "result_FT3_Quad2000.xls")  # Gets path to the .xls file in the same folder as this script
    if not os.path.isfile(workbook_path):  # Check whether the specified file exists.
        raise FileNotFoundError(f"Excel workbook not found: {workbook_path}")  # Stop if required input xls file is missing

    workbook = xlrd.open_workbook(workbook_path)
    worksheet = workbook.sheet_by_index(0)

    header_row = None  # header_row used to store the row number containing the main headers
    for row_index in range(worksheet.nrows):
        headers = [str(value).strip().casefold() for value in worksheet.row_values(row_index)]
        # a header list is created of all the values in each row
        if (
            "time and date" in headers
            and "employee number" in headers
            and "serial number" in headers
            and "record number" in headers
            and "zigbee version" in headers
            and "overall verdict" in headers
        ):    
            header_row = row_index

#Scans for key header words and if it finds them, it stores that "row" number from the excel sheet

            source_columns = {
                "time": headers.index("time and date"),
                "tester": headers.index("employee number"),
                "serial": headers.index("serial number"),
                "record": headers.index("record number"),
                "zigbee": headers.index("zigbee version"),
                "overall_verdict": headers.index("overall verdict") + 1,
            } 
            break

    if header_row is None:  # Check whether the main workbook header row was found. If it wasn't found:
        raise ValueError(
            "The first worksheet must contain the required test-result columns."
        )

    first_data_row = next(  # Find the first worksheet row containing test data.
        (
            row_index
            for row_index in range(header_row + 1, worksheet.nrows)
            if any(
                str(worksheet.cell_value(row_index, source_columns[column])).strip()
                for column in ("time", "tester", "serial")  # Visit each workbook column in the selected range.
            )
        ),
        worksheet.nrows, #The final worksheet.nrows is the fallback if no row matches.
    )

    header_rows = range(header_row, first_data_row)  # Set the rows that contain the grouped workbook headers.

    def find_header(keyword, description):  # Define a helper fn that finds a header by a text keyword.
        for candidate_row in header_rows:
            for column, value in enumerate(worksheet.row_values(candidate_row)):
                normalized_value = " ".join(str(value).split()).casefold()
                if keyword in normalized_value:
                    return candidate_row, column
        raise ValueError(f"Could not find the {description} header in the worksheet.")

    def normalized_header(value):  # Define a helper fn that makes header text case- and whitespace-insensitive.
        return " ".join(str(value).split()).casefold()

# Looks for all workbook's header rows for a cell that says "GFDI Test", and finds the position of that cell
    gfdi_header = next(
        (
            (candidate_row, column)
            for candidate_row in header_rows
            for column, value in enumerate(worksheet.row_values(candidate_row))
            if normalized_header(value) == "gfdi test"
        ), 
        None, #if does not find a match, it will return None instead of raising an error
    ) 
    if gfdi_header is None:  # if "gfdi test" was not found
        raise ValueError("Could not find the GFDI Test result header.")

# Looks for all workbook's header rows for a cell that says "step 6" and "zigbee communication", and returns position of that cell
    zigbee_section = next(
        ( 
            (candidate_row, column)
            for candidate_row in header_rows
            for column, value in enumerate(worksheet.row_values(candidate_row))
            if "step 6" in normalized_header(value) and "zigbee communication" in normalized_header(value)
        ),
        None,
    )
    if zigbee_section is None:
        raise ValueError("Could not find the Step 6 ZigBee Communication section.")

# Locates the Step 4 ABVF section in the worksheet and finds its position
    ABVF_section = next(
        (
            (candidate_row, column)
            for candidate_row in header_rows
            for column, value in enumerate(worksheet.row_values(candidate_row))
            if "step 4" in normalized_header(value) and "abnormal voltage/frequency trip" in normalized_header(value)
        ),
        None,
    )
    if ABVF_section is None: 
        raise ValueError(
            "Could not find the Step 4 Abnormal Voltage/Frequency Trip section."
        )

    ABVF_end_column = zigbee_section[1]  # Marks the beginning of the next section (Step 6 Zigbee Communication) as the right boundary for searching Step 4
    abvf_columns = {}  # Create a map of ABVF Verdict and Time-Delay columns.

    #Finds the locations of Condition 1, 2, 3 and 4 within Step 4 ABVF Section of the workbook
    for condition_number in range(1, 5):  # For Conditions 1,2,3 and 4
        condition_header = f"condition {condition_number}"
        condition = next(
            (
                (candidate_row, column)
                for candidate_row in header_rows
                for column, value in enumerate(worksheet.row_values(candidate_row))
                if ABVF_section[1] <= column < ABVF_end_column and normalized_header(value) == condition_header
            ),
            None,
        )
        if condition is None:  # If  Condition header could not be found under DC Source section
            raise ValueError( 
                f"Could not find {condition_header!r} in the Step 4 section."
            )

        condition_row, condition_column = condition  # Save the Condition header position (row, column)

        verdict_column = next(  # Find the Verdict label for this Condition.
            (
                column
                for candidate_row in range(condition_row + 1, first_data_row)  # Search the rows underneath "Condition 1,2,etc."
                for column in range(condition_column, ABVF_end_column)  # Starts at the Condition 1 column and ends at the end of step 4 ABVF
                if normalized_header(worksheet.cell_value(candidate_row, column)) == "verdict" # If the cell at (row,column) contains "verdict"
                and column < next( #Finds the column where the next Condition header begins, so for e.g. Condition 1 won't accidentally use Condition 2's Verdict
                    (   #finds next Condition's column
                        following_column
                        for following_column in range(condition_column + 1, ABVF_end_column)
                        if normalized_header(worksheet.cell_value(condition_row, following_column)).startswith("condition ")
                        #Checks whether a cell in the Condition header row starts with "condition"
                    ),
                    ABVF_end_column,  #If there isn't another Condition column, it uses ABVF_end_column as the boundary 
                )
            ),
            None,
        )
        if verdict_column is None:  # Check whether the required header or column was not found.
            raise ValueError(
                f"Could not find 'Verdict' under Step 4 {condition_header}."
            )
        #Saves the column number of Verdict column - but verdict itself has 2 columns, the colour and "Pass"
        #This saves the "Pass" value location column number 
        abvf_columns[f"ABVF Verdict {condition_number}"] = verdict_column + 1


        #Finds the column where the next Condition header begins and stores its column number here
        next_condition_column = next(
            (
                column
                for column in range(condition_column + 1, ABVF_end_column)  # Visit each workbook column in the selected range.
                if normalized_header(worksheet.cell_value(condition_row, column)).startswith("condition ")
            ),
            ABVF_end_column,
        )

        #Looks for the Time Delay (s) column and saves its value here
        delay_column = next(  # Find this Condition?s Time Delay (s) column.
            (
                column
                for candidate_row in range(condition_row + 1, first_data_row)
                for column in range(condition_column, next_condition_column)
                if normalized_header(worksheet.cell_value(candidate_row, column)) == "time delay (s)"),
            None,
        )
        if delay_column is None:
            raise ValueError(
                f"Could not find 'Time Delay (s)' under Step 4 ABVF "
                f"{condition_header}." #Specifies under which Condition e.g. "Condition 2"
            )
        abvf_columns[f"ABVF {condition_number}"] = delay_column
        """ Sample output from abvf_columns - it contains a "label":[column number]
        {
            "ABVF 1": 6,
            "ABVF 2": 10,
            "ABVF 3": 14,
            "ABVF 4": 18,
            }"""

 # Zigbee Section 
    zigbee_test_names = (  # List the six Step 6 tests used for the Zigbee summary.
        "permit joining",  # Use these headers as labels for this list
        "set sn",
        "get data",
        "get version",
        "leave device",
        "communication succeed", 
    )  
    #Finds the result column for each test in the ZigBee section and stores these column numbers for each test
    zigbee_result_columns = {}
    for test_name in zigbee_test_names: #Goes through each of the test names in ZigBee test names
        test_column = next(
            (
                column #saves into this column when the test name is found
                for candidate_row in header_rows  # Searches rows in the header_rows range
                for column, value in enumerate(worksheet.row_values(candidate_row))
                # worksheet.row_values(cadidate_row) gets all the values from the current worksheet row as a list
                # and enumerate() pairs each value with its column number, for visits those pairs one at a time
                # The loop sees: column = 0, value = "Permit joining"
                # column = 1, value = "Pass"
                # column = 2, value = "Set SN"
                # column = 3, value = "N/A"

                if column >= zigbee_section[1] and normalized_header(value) == test_name  # test name is found within the ZigBee section
            ),
            None,
        )
        if test_column is None:  # If ZigBee test header was not found
            raise ValueError(f"Could not find the Step 6 ZigBee {test_name!r} header.")
        zigbee_result_columns[test_name] = test_column + 1  # Save this Zigbee test's result column position.
        #The column is + 1 because under each test column there is a colour, then the "Pass" result - we only care about the position of the "Pass" result


    #Defines a function for finding the Verdict column of grouped tests like Auxiliary Power, Grid Detection and GFDI
    def find_group_verdict_column(group_keyword, description):  # Define a helper fn that locates a grouped test result cell.
        group_row, group_column = find_header(group_keyword, description)
        for candidate_row in range(group_row + 1, first_data_row):
            if (" ".join(str(worksheet.cell_value(candidate_row, group_column)).split()).casefold() == "verdict"
            ):
                return group_column + 1
                #group_column + 1 because the "verdict" cell is a colour, and the next column contains the actual "Pass" result
        raise ValueError(f"Could not find the {description} verdict column.")

    
    source_columns.update(
        {
            "aux_power": find_group_verdict_column("auxiliary power", "Auxiliary Power"),
            "grid_detect": find_group_verdict_column("verify grid detection function", "Grid Detection"),
            "ac_power": find_header("ac power", "AC Power")[1] + 1, #[1] + 1 moves one column over because the first col is a colour, 2nd contains the Pass/Fail result
            #find_header() returns a pair: (row number, column number)
            #the [1] returns only the column number (second item)
            "maximum_power": find_header("maximum power", "Maximum Power")[1],
            "gfdi_result": gfdi_header[1] + 1,
        }
    )

# Checks each worksheet data row to make sure it has the information needed to creates its CSV file
# Checks Time/Date and Serial number specifically 
    for row_index in range(first_data_row, worksheet.nrows): # Goes through the rows from the first data row to the end of the workheet
        source_row = worksheet.row_values(row_index)
        if not any(str(value).strip() for value in source_row):  # Check whether every examined cell is empty or whitespaced
            continue
        if ( #Checks wheter the Time and Date cell or the Serial Number cell is empty
            not str(worksheet.cell_value(row_index, source_columns["time"])).strip()
            or not str(worksheet.cell_value(row_index, source_columns["serial"])).strip()
        ): 
            raise ValueError(
                f"Worksheet row {row_index + 1} has a missing Time and Date or Serial Number; cannot determine the CSV filename."
            )

# Define the rule for N/A and blank values in standard fields in the excel sheet, NOT including MPPT/ABVF sections
# If source cell is N/A, result is also N/A in CSV
# If source cell is blank, returns ERROR in CSV cell
    def map_standard_value(value):  
        normalized_value = str(value).strip()
        if not normalized_value: #Checks if value is empty
            return "ERROR"
        if normalized_value.casefold() == "n/a":
            return "N/A"
        return value 

#Searches the workbook's header rows to find the locations of DC Source 1,2,3 and 4
    source_headers = {}  # Create a map of the DC Source header positions.
    for dc_source_number in range(1, 5):  # Repeat for DC Source numbers 1 through 4.
        source_keyword = f"dc source {dc_source_number}"  # Builds the header text to look for, such as "dc source 2"
        source_header = next(
            (
                (candidate_row, column)
                for candidate_row in header_rows
                for column, value in enumerate(worksheet.row_values(candidate_row))
                if normalized_header(value) == source_keyword
            ),
            None,
        )
        if source_header is None:  # If "DC Source" could not be found
            raise ValueError(f"Could not find the {source_keyword!r} header.")

        #source_headers[] contains the DC Source number, followed by its location (row, column)
        # 1: (4, 2),
        # 2: (4, 7),
        # 3: (4, 12),
        # 4: (4, 17),
        source_headers[dc_source_number] = source_header

    source_columns_by_dc_source = {  # Store the start column for each DC Source
        dc_source_number: source_header[1]
        for dc_source_number, source_header in source_headers.items()
    }


    MPPT_columns = {}  # Map each MPPT/PV CSV field to its corresponding columns in the workbook

    #Searches through DC Source 
    for dc_source_number in range(1, 5):  # Repeat for DC Source numbers 1 through 4.
        source_row, source_column = source_headers[dc_source_number]  # Get this DC Source's header row and starting column
        next_source_column = (  # Sets the right edge of the current DC Source's columns
            source_columns_by_dc_source[dc_source_number + 1]
            if dc_source_number < 4
            else worksheet.ncols
        )
        #Searches through Conditions under DC Source
        for condition_number in range(1, 5):  # Repeat for Condition numbers 1 through 4.
            condition_header = f"condition {condition_number}"
            condition = next(
                (
                    (candidate_row, column)
                    for candidate_row in range(source_row + 1, first_data_row)
                    for column in range(source_column, next_source_column)
                    if normalized_header(worksheet.cell_value(candidate_row, column)) == condition_header
                ), 
                None,
            )
            if condition is None:  # if Condition header could not be found in the current DC Source.
                raise ValueError(
                    f"Could not find {condition_header!r} under DC Source {dc_source_number}'."
                )
            condition_row, condition_column = condition  # Saves the (row, column) where the Condition header is found
            
            #Finds next condition
            next_condition_column = next(  # Find the next Condition boundary for the current search.
                (
                    column
                    for column in range(condition_column + 1, next_source_column)
                    if normalized_header(worksheet.cell_value(condition_row, column)).startswith("condition ")
                ),
                next_source_column, #if no next Condition, then moves onto the next Dc source column
            )


            measurement_columns = {}  # Create a map for the Current and Voltage source columns.
            for candidate_row in range(condition_row + 1, first_data_row):
                for column in range(condition_column, next_condition_column):
                    measurement_name = normalized_header(worksheet.cell_value(candidate_row, column))
                    if measurement_name in ("current (a)", "voltage (v)"):  # Check whether the cell contains Current or Voltage.
                        measurement_columns[measurement_name] = column  # Save the column where that label was found.
                        #For e.g. {"current (a)": 12, "voltage (v)": 13} --> saves the column locations where Current or Voltage was found

            for measurement_name in ("current (a)", "voltage (v)"):
                if measurement_name not in measurement_columns:
                    raise ValueError(
                        f"Could not find {measurement_name!r} under DC Source {dc_source_number}', '{condition_header}'.")
            
            #Maps the current workbook Conditions and DC source to the corresponding CSV columns. 
            #creates a shared label for the 3 CSV fields 
            csv_prefix = f"MPPT {condition_number} PV {dc_source_number}"  # Condition numbers map to MPPT numbers; DC Sources map to PV numbers.
            # i.e. Condition 4 of DC source 2 --> MPPT 4 PV 2
            MPPT_columns.update(
                {
                    f"{csv_prefix} Result": condition_column + 1,  # i.e. MPPT 4 PV 2 Result
                    f"{csv_prefix} Current": measurement_columns["current (a)"],  # i.e. MPPT 4 PV 2 Current
                    f"{csv_prefix} Voltage": measurement_columns["voltage (v)"],  # i.e. MPPT 4 PV 2 Voltage
                }
            )

#Begins the loop that processes the workbook's test-data rows, one row at a time
    for row_index in range(first_data_row, worksheet.nrows):
        date_cell = worksheet.cell(row_index, source_columns["time"])  # Read the source Time and Date cell and its Excel data type.
        serial_value = worksheet.cell_value(row_index, source_columns["serial"])  # Read the source Serial Number cell.
        if not any(str(value).strip() for value in worksheet.row_values(row_index)):  # Check whether every cell in the current row is empty
            continue

# Converts the workbook's Time and Date value into a consistent Python date/time, then formats it for the CSV
        if date_cell.ctype == xlrd.XL_CELL_DATE:
            timestamp = xlrd.xldate_as_datetime(date_cell.value, workbook.datemode)
        else:
            date_value = " ".join(str(date_cell.value).split())
            timestamp = None  # no valid time/date has been parsed yet
            for date_format in date_formats:  # Try each accepted date layout.
                try:
                    timestamp = datetime.strptime(date_value, date_format)
                    #If it works, the result is saved in timestamp, and break stops trying other formats
                    break
                except ValueError:  # format didn't match, so continue tries the next one
                    continue
            if timestamp is None: #if none of the formats work, stays None, and raises an error
                raise ValueError(
                    f"Invalid Time and Date value on worksheet row {row_index + 1}: {date_value!r}")

        #Formats the CSV date and time with a four-digit year.
        date_value = timestamp.strftime("%Y/%m/%d %H:%M:%S")

        # Converts serial number to text
        if isinstance(serial_value, float) and serial_value.is_integer(): #Checks whether value is a decimal number
            serial_number = str(int(serial_value))  # removes the decimal and converts number to integer as text
        else:
            serial_number = str(serial_value).strip()


        #Builds filename of the CSV and checks whether the serial number in the filename matches the serial number in the original source file
        filename = f"FT3_{serial_number}_{timestamp:%Y_%m_%d_%H-%M-%S}.csv"
        filename_serial = filename[len("FT3_") :].rsplit(f"_{timestamp:%Y_%m_%d_%H-%M-%S}.csv", 1)[0]
        if filename_serial != serial_number:
            raise ValueError(
                f"Entry serial number {serial_number!r} does not match the serial number in file name {filename!r}."
            )

        #Creates a new, blank CSV data row with one empty string for every CSV column
        # Creates something like: ["", "", "", "", "", "", "", "", "", ""]
        row = [""] * len(csv_headers) 


        # Pairing CSV column (header) name entries with the values for the source Excel workbook
        #row_index starts from the first data row in the worksheet
        mapped_values = {
            "Time and Date": date_value,
            "Tester": map_standard_value(worksheet.cell_value(row_index, source_columns["tester"])),
            "Serial Number": map_standard_value(serial_number),
            "Overall Verdict": map_standard_value(worksheet.cell_value(row_index, source_columns["overall_verdict"])),
            "FPGA Version": map_standard_value(worksheet.cell_value(row_index, source_columns["record"])),
            "Zigbee Version": map_standard_value(worksheet.cell_value(row_index, source_columns["zigbee"])),
            "Aux Power Result": map_standard_value(worksheet.cell_value(row_index, source_columns["aux_power"])),
            "Grid Detect Result": map_standard_value(worksheet.cell_value(row_index, source_columns["grid_detect"])),
            "AC Power": map_standard_value(worksheet.cell_value(row_index, source_columns["ac_power"])),
            "Maximum Power": map_standard_value(worksheet.cell_value(row_index, source_columns["maximum_power"])),
            "GFDI Result": map_standard_value(worksheet.cell_value(row_index, source_columns["gfdi_result"])),

            #Adds the 4 ABVF time delay values to the mapped_values{}
            #loops through ABVF_columns, which contains labels and column numbers i.e. [ABVF 1: 6]
            **{
                csv_header: ( #If the cell is empty, it populates it with 'N/A', otherwise it keeps its original value
                    "N/A"
                    if not str(value).strip()
                    else value
                )
                for csv_header, source_column in abvf_columns.items() #loop through the entries in abvf_columns, and skip the verdict entries
                if not csv_header.startswith("ABVF Verdict ") #skips verdict fields, this loop is only for delay values
                for value in (worksheet.cell_value(row_index, source_column),)
            },
        }

        #Reads the 4 ABVF Condition verdicts then valculates the overall "ABVF Result"
        abvf_verdicts = [
            (
                str(worksheet.cell_value(row_index, abvf_columns[f"ABVF Verdict {condition_number}"])).strip().casefold() or "n/a" #reads the cell
            )
            for condition_number in range(1, 5)]  #For each Condition 1-4, reads the cells from that row
        
        if all(verdict == "n/a" for verdict in abvf_verdicts): #If all N/A --> ABVF Result is N/A
            mapped_values["ABVF Result"] = "N/A" 
        elif all(verdict == "pass" for verdict in abvf_verdicts): #If all pass --> ABVF Result is Pass
            mapped_values["ABVF Result"] = "Pass"
        else:
            mapped_values["ABVF Result"] = "Fail" #else ABVF Result = Fail


        #Adds all the MPPT/PV values for the current row to the mapped_values{}
        mapped_values.update(
            {
                column: (
                    "N/A"
                    if not str(value).strip() or str(value).strip().casefold() == "n/a"  # if blank, then corresponds to N/A in the CSV
                    else value
                )
                for column, source_column in MPPT_columns.items() # i.e. MPPT 4 PV 2 : [column number]
                for value in (worksheet.cell_value(row_index, source_column),)
            }
        ) 

        # Gets "MPPT Result" entry in CSV file
        condition_results = [
            (
                str(worksheet.cell_value(row_index, source_column)).strip().casefold() or "n/a"
            )
            for column, source_column in MPPT_columns.items() #For each entry in MPPT list
            if column.endswith(" Result") #Gets all the entries with "Result"
        ]
        if all(result == "n/a" for result in condition_results): #If all MPPT Condition results were "N/A"
            mapped_values["MPPT Result"] = "N/A"  # MPPT Results = N/A
        elif all(result == "pass" for result in condition_results): #If all MPPT Conditions were "Pass"
            mapped_values["MPPT Result"] = "Pass"  #MPPT Results = Pass
        else:
            mapped_values["MPPT Result"] = "Fail"  # Else MPPT Results = Fail

        # Gets overall ZigBee Results, based on it's Passes and Fails under its categories
        zigbee_results = [
            (
                str(worksheet.cell_value(row_index, source_column)).strip().casefold() or "n/a"
            )
            for source_column in zigbee_result_columns.values()  # Reads all columns under Zigbee results
        ]
        if any(not str(worksheet.cell_value(row_index, source_column)).strip()
               for source_column in zigbee_result_columns.values()):
            mapped_values["Zigbee Result"] = "ERROR"  # If any values are blank, set ZigBee result = ERROR
        elif all(result == "n/a" for result in zigbee_results):  # If all results are N/A, ZigBee Result = N/A
            mapped_values["Zigbee Result"] = "N/A"
        elif all(result == "pass" for result in zigbee_results):  # If all Pass, Zigbee Result = Pass
            mapped_values["Zigbee Result"] = "Pass" 
        else:
            mapped_values["Zigbee Result"] = "Fail"  # Else Zigbee Result = Fail

        #Puts values for the current test record into a CSV row, then checks wheter that record's CSV file already exists
        for column, value in mapped_values.items():
            row[csv_headers.index(column)] = value 

        # Build the full output path for this record's CSV file
        csv_path = os.path.join(output_results_dir, filename)

        if os.path.isfile(csv_path):  # Check whether the specified file exists.
            with open(csv_path, newline="", encoding="utf-8") as csv_file: 
                existing_rows = list(csv.reader(csv_file))

            if not existing_rows:  # If the existing CSV file has no rows
                existing_rows = [csv_headers]
            elif existing_rows[0] == ["Test Time"] + csv_headers[1:]:
                existing_rows[0] = csv_headers
            elif existing_rows[0] != csv_headers:
                raise ValueError(f"Unexpected CSV header in {csv_path!r}.")


            #Checks the data rows already in the CSV file to make sure they belong to the same serial number and test date/time as the file's name
            matching_row = None
            for existing_row in existing_rows[1:]:  # Loops through the CSV rows after the header. Skips the first row containing the names
                if len(existing_row) != len(csv_headers):
                    raise ValueError(f"Unexpected CSV row width in {csv_path!r}.")
                existing_serial = existing_row[csv_headers.index("Serial Number")].strip()  # Read the serial number field from that CSV row
                if existing_serial != filename_serial:  # Check whether the serial number disagrees with the filename SN
                    raise ValueError(
                        f"Entry serial number {existing_serial!r} does not match the serial number in file name {filename!r}."
                    )

                #Checks the Time and Date value in an existing CSV data row
                existing_time = existing_row[csv_headers.index("Time and Date")]
                existing_timestamp = None
                for date_format in date_formats:  # Try each accepted date layout.
                    try: 
                        existing_timestamp = datetime.strptime(existing_time, date_format)
                        break
                    except ValueError:
                        continue
                existing_date_only = existing_time == timestamp.strftime("%Y/%m/%d")
                if existing_timestamp is None or (
                    existing_timestamp != timestamp and not existing_date_only
                ):
                    raise ValueError(
                        f"Entry Time and Date {existing_time!r} does not match the date and time in file name {filename!r}."
                    )
                if existing_serial == serial_number:
                    matching_row = existing_row
                    break

            if matching_row is None:  # Earlier search didn't find and existing row for this record
                existing_rows.append(row)  # Adds the newly prepared data row
            else:
                for column, value in mapped_values.items():  # Visit each column together with its cell value.
                    #The loop goes through each CSV column name and its new value, finds that column's position and replaces the value in that position.
                    #This updates the existing row rather than creating a duplicate
                    matching_row[csv_headers.index(column)] = value

            with open(csv_path, "w", newline="", encoding="utf-8") as csv_file:
                csv.writer(csv_file).writerows(existing_rows)  # Write all rows to the CSV using proper CSV formatting.
        
        else: #If the CSV file doesnt exist yet
            #Writes into the new CSV file
            with open(csv_path, "w", newline="", encoding="utf-8") as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(csv_headers)  # Write CSV headers into the CSV file
                writer.writerow(row)  # Writes current test record into the next row, putting each value in its corresponding column


create_result_csv_files()
