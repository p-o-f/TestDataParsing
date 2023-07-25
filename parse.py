import pandas as pd
import os
#import openpyxl
#from openpyxl import Workbook


def delete_file(path): # Pass in the extension as well, such as example.txt or example.xlsx, not just example
    if os.path.exists(path):
        os.remove(path)
    else:
        print("File that was attempted to be deleted does not exist")


def sheet_exists(excel_filename, sheetname): # Returns boolean if a sheet exists in a given xlsx file
    path = excel_filename + ".xlsx"
    try:
        sheets = pd.ExcelFile(path).sheet_names
        if sheetname in sheets:
            return True
        return False
    except:
        return False

def get_length_of_longest_list(lst):
    return max(len(x) for x in lst)


def xlsx(dataframe, excel_filename, sheetname): # Exports a dataframe to xlsx; either make a new xlsx file and a new sheet for the dataframe, or just make a new sheet into existing xlsx and add the dataframe there   
    path = excel_filename + ".xlsx"
    try: # First, try to add a new sheet to the current xlsx file if it already exists
        with pd.ExcelWriter(path, mode="a") as writer: 
            dataframe.to_excel(writer, sheet_name=sheetname)
    except: # If no xlsx file has yet been created, then make a new one
        dataframe.to_excel(path, sheet_name = sheetname)


#TODO github commit/publish when done... also publish the example txt file along with this code
#TODO compile everything into one .py file so import statements or pandas install etc is unneeded
#TODO implement main loop with prompts IE type stop at any time to stop, type back at any time to go back, more print statements to give more feedback about what was done.. etc.


def abort(userinput): 
    if (userinput == "e"): # Exit
        exit()
    elif (userinput == "r"): # Restart to re-enter parameters
        start()
    else: # Continue start() function as normal
        return()


def parse_by_number(txt_filename, query):
    data = [] # A list-of-lists that will collect data and be used to initialize a dataframe when ready
    numOfColumns = 0 # Number of columns for the future dataframe
    columnBase = ["Number", "Site", "Result", "Test Name"] # Every test number has at least these rows, but some have more
    check = 0 # Flag to see if test number is found/the data is then successfully exported to xlsx
    
    with open(txt_filename + ".txt") as reader: # Open the file for reading
        for line in reader: # Process the file line by line
            if line.isspace() == False: # Ignore completely blank lines
                splitline = line.split() # Split each individual line into a list with whitespace as the delimiter
                if (query == splitline[0]): # The first element of the line should be the test number / the forth element of the line should be the test name
                    if check == 0:
                        check = 1 # At least one instance of the data has been found
                    if numOfColumns == 0:
                        numOfColumns = len(splitline) # Every instance of a test number should have the same number of columns
                    data.append(splitline) # Each element in data is going to be a row in the future dataframe
                    
    # Leave the loop and file reader; no longer need to process the file, all relevant data is now stored in the data list
    extraColumns = numOfColumns - 4
    if (extraColumns > 0): # This means it's going to be more columns, something like: Number  Site Result   Test Name       Pin      Channel Low            Measured       High..... etc
        for i in range(extraColumns):
            columnBase.append(i+1) # The extra columns are going to be titled in the dataframe from 1 to N where N is the number of columns beyond the standard 4th column of "Test Name"
    
    df = pd.DataFrame(data, columns=columnBase) # Generate the dataframe to export to excel
    return df, check


def parse_by_name(txt_filename, query):
    data = [] # A list-of-lists that will collect data and be used to initialize a dataframe when ready
    numOfColumns = 0 # Number of columns for the future dataframe
    columnBase = ["Number", "Site", "Result", "Test Name"] # Every test number has at least these rows, but some have more
    check = 0 # Flag to see if test number is found/the data is then successfully exported to xlsx
    with open(txt_filename + ".txt") as reader: # Open the file for reading
        for line in reader: # Process the file line by line
            if line.isspace() == False: # Ignore completely blank lines
                splitline = line.split() # Split each individual line into a list with whitespace as the delimiter
                if (query in splitline):
                    if check == 0:
                        check = 1 # At least one instance of the data has been found
                    data.append(splitline) # Each element in data is going to be a row in the future dataframe
                    
	# Leave the loop and file reader; no longer need to process the file, all relevant data is now stored in the data list
    numOfColumns = get_length_of_longest_list(data)
    extraColumns = numOfColumns - 4
    if (extraColumns > 0): # This means it's going to be more columns, something like: Number  Site Result   Test Name       Pin      Channel Low            Measured       High..... etc
        for i in range(extraColumns):
            columnBase.append(i+1) # The extra columns are going to be titled in the dataframe from 1 to N where N is the number of columns beyond the standard 4th column of "Test Name"
            
    df = pd.DataFrame(data, columns=columnBase) # Generate the dataframe to export to excel
    df.fillna("NA")
    print(df)      
    
#parse_by_name("char", "703701") TODO look at this, something like this is viable but probably slightly more resource intensive than parse_by_number? so probably keep both methods?? but examine later
# TODO ^ Note the above should work with infinite number of columns in any scenario as long as the base columns: number  site  result  testname, are there 
parse_by_name("char", "por_char_F_min")

def start(xlsx_path=None, search_by_name=False):
    
    print("\n.txt to .xlsx script. Follow the provided prompts. At any point, type e to exit or r to re-enter parameters.")
    
    while True: # Prevent searching for a blank file
        txt_filename = input("\nEnter the .txt filename to search: ")
        abort(txt_filename)
        if txt_filename != "" and txt_filename.isspace() == False:
            break
        
    if (xlsx_path is None): # Prompt for a xlsx filename if not recursively adding another sheet to the same xlsx file
        while True: # Prevent exporting to a blank file
            xlsx_path = input("\nEnter the .xlsx filename to export to: ")
            abort(xlsx_path)
            if xlsx_path != "" and xlsx_path.isspace() == False:
                break
    
    while True: # Do not allow exporting to same sheet; blank/whitespace will still pass this and continue to use the default sheet name
        xlsx_sheet = input("\n(Optional, may leave this blank) Enter the sheet name to export to: ")
        abort(xlsx_sheet)
        if sheet_exists(xlsx_path, xlsx_sheet) == False: 
            break
    
    while True: # Prevent searching for nothing or for whitespace
        if (search_by_name == False):
            query = input("\nEnter the term (Test NUMBER) you would like to search for: ")
            abort(query)
        else: 
            query = input("\nEnter the term (Test NAME) you would like to search for: ")
            abort(query)
        if query != "" and query.isspace() == False:
            break
    
    if (xlsx_sheet.isspace()) or (xlsx_sheet == ""): # Default sheet name/fail safe
        xlsx_sheet = query # Set the default sheet name to be the test number
        
    try:
    
        if search_by_name == False: # Procedure if searching by test NUMBER
            df, check = parse_by_number(txt_filename, query)
            
        
        if (check != 1):
            retry = input("Item not found. Type y to run again or anything else to exit: ")
            if (retry == "y"):
                start() # Try again
            return() # Exit
        
        else:
            print("\n Found the following data and exported it to excel: \n")
            print(df)
            xlsx(df, xlsx_path, xlsx_sheet)
            
            add_new_sheet = input("\nWould you like to export another test number into the same excel file, into a new sheet? Type y to do so or anything else to exit: ")
            if (add_new_sheet == "y"): 
                start(xlsx_path) # Recursively add another test number into the same xlsx file, but in a new sheet 
                
            return() # Exit
        
    except:
        print("Error. One of the supplied parameters may have been incorrect? Retrying...\n")
        start() # Try again




#TODO add functionality to prompt user to do another test number, and make a new sheet for that test number unless it's the same one that just got asked (then state that the user just 
# asked for that test number and reprompt)... might want to have a global test numbers list[] variable for this

# add functionality to get the SITE SORT BIN and Site failed tests/executed tests
# ^^^ Unneeded for now




#TODO add functionality to be able to input a Test Name, not just a test number... approach: this will probably need multiple dataframes where you assign each value of the
# data list into a dataframe depending on the number of elements in it and then put all the dataframes into a single sheet or multiple sheets?





#TODO add if there is an error, delete any existing files which were created

#TODO ask if it's a good idea to mass dump all the test numbers into sheets with their respective numbers into one giant excel file?


#start()





# https://stackoverflow.com/questions/13784192/creating-an-empty-pandas-dataframe-and-then-filling-it
# https://www.geeksforgeeks.org/different-ways-to-create-pandas-dataframe/    used method 2 


# https://sparkbyexamples.com/pandas/pandas-write-to-excel-with-examples/?expand_article=1
# https://stackoverflow.com/questions/42370977/how-to-save-a-new-sheet-in-an-existing-excel-file-using-pandas
# https://stackoverflow.com/questions/17977540/pandas-looking-up-the-list-of-sheets-in-an-excel-file



