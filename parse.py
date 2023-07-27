import pandas as pd
import os



def delete_files(directory = os.getcwd()): 
    """
    Deletes all the .xlsx files created within the specified directory.

    This will always default to the current directory unless a parameter is specified.

    Parameters:
    directory (string, Optional): Directory where files should be deleted

    Returns:
    None

    """
    
    files = os.listdir(directory)
    for file in files:
        if file.endswith(".xlsx"):
            os.remove(os.path.join(directory, file))
           


def sheet_exists(excel_filename, sheetname): 
    """
    Returns TRUE if a sheet exists and FALSE if not.

    Returns a boolean value for if a sheet with a certain name exists in a given xlsx file.

    Parameters:
    excel_filename (string): The excel/xlsx filename without ".xlsx"
    sheetname (string): The name of the sheet that is being verified
    
    Returns:
    bool: True if the sheet exists, False if not

    """
    
    path = excel_filename + ".xlsx"
    try:
        sheets = pd.ExcelFile(path).sheet_names
        if sheetname in sheets:
            return True
        return False
    except:
        return False



def get_length_of_longest_list(lst):
    """Returns (int) the length of the longest sublist contained in lst; lst should be a list-of-lists."""
    
    return max(len(x) for x in lst)



def xlsx(dataframe, excel_filename, sheetname): 
    """
    Exports a Pandas DataFrame to an existing or new .xlsx file.

    Either makes both a new .xlsx file and a new sheet for the dataframe to go into, or only makes a new sheet into an existing .xlsx and adds the dataframe there.

    Parameters:
    dataframe (DataFrame): a single Pandas DataFrame which should be prepared to be exported to excel
    excel_filename (string): the excel/xlsx filename without ".xlsx"
    sheetname (string): the name of the sheet in the excel file that will contain the DataFrame
    
    Returns:
    None

    """
     
    path = excel_filename + ".xlsx"
    try: # First, try to add a new sheet to the current xlsx file if it already exists
        with pd.ExcelWriter(path, mode="a") as writer: 
            dataframe.to_excel(writer, sheet_name=sheetname)
    except: # If no xlsx file has yet been created, then make a new one
        dataframe.to_excel(path, sheet_name = sheetname)



def abort(userinput):
    """Helper function for the main start() driver function.""" 
    
    if (userinput == "e"): # Exit
        exit()
    elif (userinput == "r"): # Restart to re-enter parameters
        start()
    else: # Continue start() function as normal
        return()



def look_for_data(txt_filename, query, deleteNoneType = True):
    """
    Exports a Pandas DataFrame to an existing or new .xlsx file.

    Either makes both a new .xlsx file and a new sheet for the dataframe to go into, or only makes a new sheet into an existing .xlsx and adds the dataframe there.

    Parameters:
    txt_filename (string): the .txt filename without ".txt"
    query (string): CASE SENSITIVE... the item to search (verified use cases: any test number, any test name, "FAIL" for every failing test, "PASS" for every passing test)
    deleteNoneType (bool, Optional): if disabled, entries where data does not exist will hold the string value "None" in the final excel file
    
    Returns:
    int: 0 if no data is found, 1 if any instance of relevant data is found
    DataFrame: A dataframe that will be exported to excel containing all relevant data that was searched for

    """
    
    data = [] # A list-of-lists that will collect data and be used to initialize a large dataframe when ready
    numOfColumns = 0 # Number of columns for the future dataframe
    columnBase = ["Number", "Site", "Result", "Test Name"] # Every test number has at least these rows, but some have more
    check = 0 # Flag to see if test number is found/the data is then successfully exported to xlsx
    with open(txt_filename + ".txt") as reader: # Open the file for reading
        for line in reader: # Process the file line by line
            if (line.isspace() is False): # Ignore completely blank lines
                splitline = line.split() # Split each individual line into a list with whitespace as the delimiter
                if (query in splitline): # Check if an EXACT match for query is in the line (i.e. 10000 should return false if checking for 100)
                    if (check == 0):
                        check = 1 # Update check to reflect that at least one instance of the data has been found
                    data.append(splitline) # Each list element in data is going to be a row in the future dataframe
                    
	# Leave the loop and file reader; no longer need to process the file, all relevant data is now stored in the data list
    numOfColumns = get_length_of_longest_list(data) # Need as many columns as the length of the longest list
    extraColumns = numOfColumns - 4 # Number of columns after the columBase of: "Number   Site   Result   TestName"
    if (extraColumns > 0): # This means it's going to be more columns, something like: Number  Site Result   TestName     Pin     Channel Low     Measured   High..... etc
        for i in range(extraColumns):
            columnBase.append(i+1) # The extra columns are going to be titled in the dataframe from 1 to N where N is the number of columns beyond the columnBase of 4 columns
                                   # The reason for this is because it is impossible to know for sure what data these extra columns will contain, as it's variable
    
    # Prepare and return the dataframe and check flag
    df = pd.DataFrame(data, columns=columnBase) # Generate the dataframe to export to excel
    if (deleteNoneType is True): # Replace the entries of "None" to be blanks
        df.fillna("", inplace=True)
    return df, check



def start(xlsx_path=None):
    """Driver function; prints necessary prompts to terminal and collects input. The xlsx_path parameter is used if recursively called (export to same xlsx file)."""
    
    linebreak = "--------------------------------------------------------------------------------------------------------------" # Just for formatting to terminal
    print("\nFollow the provided prompts. At any point, type e to exit or r to re-enter parameters.")
    print(linebreak)
    
    while True: # Prevent searching for a blank file
        txt_filename = input("\nEnter the .txt filename to search: ")
        abort(txt_filename)
        if txt_filename != "" and txt_filename.isspace() is False:
            break
    print(linebreak)
        
    if (xlsx_path is None): # Prompt for a xlsx filename if not recursively adding another sheet to the same xlsx file
        while True: # Prevent exporting to a blank file
            xlsx_path = input("\nEnter the .xlsx filename to export to: ")
            abort(xlsx_path)
            if xlsx_path != "" and xlsx_path.isspace() is False:
                break
    print(linebreak)
    
    while True: # Do not allow exporting to same sheet; blank/whitespace will still pass this and continue to use the default sheet name
        xlsx_sheet = input("\n(Optional, may leave this blank) Enter the sheet name to export to: ")
        abort(xlsx_sheet)
        if sheet_exists(xlsx_path, xlsx_sheet) is False: 
            break
    print(linebreak)
    
    while True: # Prevent searching for nothing or for whitespace
        query = input("\nEnter the term you would like to search for (test number, test name, PASS, FAIL): ")
        abort(query)
        if query != "" and query.isspace() is False:
            break
    print(linebreak)
    
    if (xlsx_sheet.isspace()) or (xlsx_sheet == ""): # Default sheet name/fail safe
        xlsx_sheet = query # Set the default sheet name to be the item being searched for
        
    try:
        df, check = look_for_data(txt_filename, query) # If check = 1, data has been found successfully
        if (check != 1):
            retry = input("Item not found. Type y to run again or anything else to exit: ")
            if (retry == "y"):
                start() # Try again
            return() # Exit
        else:
            print(linebreak)
            print("\n Found the following data and exported it to excel: \n")
            print(df)
            xlsx(df, xlsx_path, xlsx_sheet)
            
            add_more = input("\nWould you like to export anything else? Type y to do so or anything else to exit: ")
            if (add_more == "y"): 
                same_sheet = input("Would you like to export to the same excel file in a new sheet? Type y to do so or anything else to change the target excel sheet to export to.")
                if (same_sheet == "y"):
                    start(xlsx_path) # Recursively add into the same xlsx file, but in a new sheet 
                else:
                    start() # Reprompt for xlsx file target; add to a specified xlsx file
            return() # Exit
    except:
        print("Error. One of the supplied parameters may have been incorrect? Retrying...\n")
        delete_files() # Delete any .xlsx files that were created in the current directory
        start() # Try again



print("Data Parsing Script: .txt to .xlsx")
print("Please have relevant .txt files in the same directory as this script, or this script will not be able to read them!")
print("If you want to add sheets to an existing .xlsx file or multiple existing .xlsx files, have all relevant .xlsx files in the same directory as well.")
start()

