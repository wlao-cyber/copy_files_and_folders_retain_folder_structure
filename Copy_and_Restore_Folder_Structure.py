""" This copy program restores the full folder structure
after the mapped drive letter. """

# IMPORTING PACKAGES/MODULES
# os to read, direct, and create file/folder paths
# subprocess to run Windows commands
# csv to generate csv file
# time to print out a unique errors.csv named by a unique current date & time that doesn't overwrite older ones
# Use regular expressions to check source path starting with mapped drive letter
# datetime to calculate the execution time
# traceback to show any program crash errors.
import os
import subprocess
import csv
import time
import re
import datetime
import traceback

# start time of script
start_time = datetime.datetime.now()

# win32com.client (part of pywin32 package) to use Windows' File Explorer
# xlrd to read from Excel file
import win32com.client
import xlrd

def variables():
    # set global variables
    global fso, file_location, workbook, sheet, invalid_paths_rows, \
    empty_rows, missing_entries_rows, \
    size_mismatch_rows, copy_errors, unmapped_src_rows
    
    # use Win32 API to get the folder size later.
    # set file system object using the Dispatch command
    # Provides access to a computer's file system
    fso = win32com.client.Dispatch("Scripting.FileSystemObject")

    # Specify the Excel file to read from. The long file path is unnecessary if the script is in the same directory,
    # but specifying the long file path of the excel file gives you the option to run the script from anywhere. 
    file_location = r"Copy_Files_and_Folders_Retain_Structure.xlsx"

    workbook = xlrd.open_workbook(file_location)

    # This opens the first sheet (0th index) in the Excel File
    # If you know the name of the sheet, you can also open by the sheet name "Sheet1" in this case.
        #sheet = workbook.sheet_by_name("Sheet1")
    sheet = workbook.sheet_by_index(0)

    # set variable out of loop to keep track of invalid paths as well as empty excel rows.
        # Use a list to keep adding rows to.
        # Then shoot out the result at the end, detailing which rows are empty or are missing entries.
        # mainly using lists because lists retain order.
    unmapped_src_rows = []
    empty_rows = []
    missing_entries_rows = []
    invalid_paths_rows = []
    size_mismatch_rows = 0

    # set global list of dictionaries objects that will keep growing as size mismatches are detected.
    copy_errors = []

def read_Excel():
    # set global for size_mismatch_rows again so other functions like this can better identify
    # Otherwise, python may assume local variable is referenced before being assigned
    global size_mismatch_rows
    
    # Read through the spreadsheet.
    # Ignore the first row because it has the header. Set the for loop to start with the second Excel row
    # Use the cell values indicated in the Excel document to copy the files/folders to their destination folders
    for row in range(1, sheet.nrows):
        
        # the argumeent next to row is the column index 0th and 1st, which means the first and second Excel columns.
        # Use .strip() method so that leading and trailing whitespace will be ignored in original and destination path
            # problem is that long numbers inputted into the Excel file will be interpreted as a float.
            # Thus, the .strip() method would fail on a type that is not a string
            # safer to convert the src_path and dest_path type to str preemptively, so strip will work
        src_path = str(sheet.cell_value(row, 0)).strip()
        dest_path = str(sheet.cell_value(row, 1)).strip()

        # Check where both src path and dest path are not null after being stripped
            # This is to exclude cells that have only spaces but no actual values
        if src_path != '' and dest_path != '':

            # set drive letter pattern to use later for conditions
            drive_letter = r"^[a-zA-Z]:\\"

            # Need to check if dest path exists. If not, create the dest path
                # If dest path is not a full path (ex. just alphanumeric chars), then a folder will be created where this program is
                # To avoid this, set the condition to check if the destination is an abs path before creating folders
            # Only create the dest path if:
                # the src path exists, if there's actually something valid to copy over
                # destination path specified is a legitimate path to create
                    # os.path.isabs() can be used to see if path begins with backslash after drive letter in Windows
                # dest_path does NOT already exist.
            if os.path.exists(src_path) and os.path.isabs(dest_path) and os.path.exists(dest_path)==False:

                # However, os.path.isabs() does not check to see if path will be valid
                    # leading and trailing spaces are illegal in Windows folder/file names but return True 
                    # to os.path.isabs() as long as the path begins with backslash after a drive letter or long path
                # Solution: Remove leading and trailing whitespace in dest path set by user input (if any)
                    # Ex. remove spaces in folder like C:\   Users\ Ocelot  \Demo \
                    # split dest path by its backslashes to separate all folders in a list
                    # Use list comprehension to quickly strip leading and trailing whitespace in the folder names
                        # is like in place redefining of the list items
                    # join all list items by backslash
                        # This still returns long UNC file paths to original because list nulls from stripping are
                        # joined by '\\' which in turn restores the original '\\\\' that starts the long UNC paths 
                split_path = dest_path.split('\\')
                strip_whitespace = [i.strip() for i in split_path]
                dest_path = '\\'.join(strip_whitespace)

                # ONLY create the stripped full path if it doesn't already exist.
                # Sometimes, you can have different user input that reevaluates to the same path after being stripped.
                    # Ex. of two folders that will reevaluate to same path after stripping
                        # D:\  whitespace folder \trailing \ leading
                        # D:\  whitespace folder  \            trailing \          leading
                # Otherwise, will get a traceback error.
                if os.path.exists(dest_path)==False:
                    
                    # os.makedirs() will create a path whether or not the string ends with a backslash \
                    # also creates nonexisting intermediate paths.
                    print("Creating nonexisting destination path: " + dest_path)
                    os.makedirs(dest_path)

            # Robocopy follows different syntax for folders and for files
            # First check for when source path is a dir and destination path is a dir to use robocopy format for copying folders
            # The following if statement only runs IF both src and dest paths exist
                # from the above if statement, the dest path should be created if didn't exist originally
            # include condition to check if src_path begins with drive letter
                # re.match('pattern', 'string') checks only if a string's pattern matches at the start.
                    # bool to return True or False
                # regular expression for start of a mapped drive letter, lowercase or uppercase:
            if os.path.isdir(src_path) and os.path.isdir(dest_path) and bool(re.match(drive_letter, src_path)):
                # os.path.abspath() to gets the full path regardless of whether there's a \ at the end or not.
                # set both src_path and dest_path as abspaths so that if there's a backslash at the end, it gets ommitted.
                src_path = os.path.abspath(src_path)
                dest_path = os.path.abspath(dest_path)

                # Robocopy doesn't copy the base folder over to destination path.
                # Therefore, add the base folder to the end of the destination path.

                # Use os.path.basename() to get the base folder or file name in the source path
                # Then add the basename to the destination folder to include the src's base folder, making the final destination

                # Easy solution that strips only the first 3 characters and adds the rest to destination path
                # better than using the .find() method to look for the base folder because you can have folders with the same name.
                folder_structure = src_path[3::]                
                
##                src_base_folder = os.path.basename(src_path)
                
                final_dest = dest_path + '\\' + folder_structure
                
                print("Source path: " + src_path)
                print("Destination path: " + dest_path)
                
                # Use subprocess module to run Robocopy from the source path to the final destination
                    # Robocopy syntax for folders: Robocopy "src folder" "dest folder" [flags]
                    # /COPY:DAT copies all file properties for data, attribute, and timestamps
                    # /E flag copies all subdirectories in source path, including empty folders
                subprocess.run(["Robocopy", src_path, final_dest, "/COPY:DAT", "/E"])

                # get the size of the source path and the final dest path (just the copied base folder)
                # point Windows API's FileSystemObject to the folder path to get the sizes in bytes
                src_fldr = fso.GetFolder(src_path)
                dest_fldr = fso.GetFolder(final_dest)

                scrc_fldr_size = src_fldr.Size
                dest_fldr_size = dest_fldr.Size
                
                print("Source Folder Size: " + str(scrc_fldr_size))
                print("Destination Folder Size: " + str(dest_fldr_size))
                print()
                # If size mismatch detected, compile dictionary terms into the copy_errors list defined globally
                    # Headers will be 'Source Path', 'Destination Path', 'Source Size', 'Destination Size'
                    # Set the values tied to these dictionary items
                        # dictionary object can have combination of variables, numbers, and strings. (i.e., mix of strings and numbers)
                if scrc_fldr_size != dest_fldr_size:
                    copy_errors.append({'Source Path' : src_path, 'Destination Path' : final_dest, \
                                        'Source Size' : scrc_fldr_size, 'Destination Size' : dest_fldr_size})
                    # for each size mismatch, add to counter
                    size_mismatch_rows += 1

            # Check for when source path is a file and destination path is a dir to use different robocopy format
                # similar to before, use os.path.basename() grab just the base file name just for the robocopy command.
                # Does NOT work if there is a trailing backslash at the end of file name.
                    # This wouldn't make sense anyway and would be placed under invalid paths
            # just like before, use regular expression to check if src_path begins with drive letter
            elif os.path.isfile(src_path) and os.path.isdir(dest_path) and bool(re.match(drive_letter, src_path)):
                src_path = os.path.abspath(src_path)
                dest_path = os.path.abspath(dest_path)
                
                src_base_file = os.path.basename(src_path)

                print("Source path: " + src_path)
                print("Destination path: " + dest_path)

                # Cannot just strip only the first 3 characters and add the rest to destination path
                # Otherwise, would include the base file as a redundant folder
                # use src_base_file to find the index where base file is located in the src_path string
                    # Then set that take only the part of the src_path string after the first 3 characters [drive letter]
                    # but before the base name
                    # This will be the folder structure after the mapped drive but before the base file
                #.find() can be used to find the first point where the src_base_file is in the full source path
                    # Unlikely, but you can have a folder named exactly as the src_base_file higher up (as in the folder above)
                    # This will cause the index to be found where the folder with the same name as the file
                # end_index = src_path.find(src_base_file)
                # Use .rfind() to return the highest index where the substring is, so as close to the end of the source path.
                    # This will avoid cutting off the folder structure where a higher level folder exists with same named
                    # as the src base file.
                end_index = src_path.rfind(src_base_file)
                folder_structure = src_path[3:end_index]
                
##                print("Restored source file's folder structure: " + folder_structure)
                
##                src_base_folder = os.path.basename(src_path)
                
                final_dest = dest_path + '\\' + folder_structure

                # command syntax for files: robocopy "src folder" "dest folder" file.txt [flags]
                subprocess.run(["Robocopy", os.path.dirname(src_path), final_dest, src_base_file, "/COPY:DAT"])

                # point FileSystemObject to the file path to get the sizes in bytes
                src_file = fso.GetFile(src_path)
                dest_file = fso.GetFile(final_dest + '\\' + src_base_file)

                src_file_size = src_file.Size
                dest_file_size = dest_file.Size
                
                print("Source File Size: " + str(src_file_size))
                print("Destination File Size: " + str(dest_file_size))
                print()
                # Just like before, if size mismatch detected, append dictionary terms into the copy_errors list
                if src_file_size != dest_file_size:
                    copy_errors.append({'Source Path' : src_path, 'Destination Path' : dest_path, \
                                        'Source Size' : src_file_size, 'Destination Size' : dest_file_size})
                    size_mismatch_rows += 1

            # Add to tracker if src path exists but does not start with a mapped drive
            # Append the Excel row into the master list for unmapped paths
                # Excel row is "row+1" because xlrd interprets first row as 0th row.
            elif os.path.exists(src_path) and bool(re.match(drive_letter, src_path))==False:
                unmapped_src_rows.append(row+1)

            # For all cases where, the row values do NOT match the above conditions of:
                # Copy valid folder path to valid folder path
                # Copy valid file path to valid folder path
            # Append the Excel row into the master list for invalid paths
                # Excel row is "row+1" because xlrd interprets first row as 0th row.
            else:
                invalid_paths_rows.append(row+1)
        
        # The elif statement will check when both src and dest are missing values
        # append the Excel row number to the master empty_rows list 
        elif src_path == '' and dest_path == '':
            empty_rows.append(row+1)

        # The else statement will check if either the srcs or dest are missing values
        # append the Excel row number to the missing_entries_rows list variable.
        else:
            missing_entries_rows.append(row+1)

# Calculate the difference from start to end time and then convert to str for readable format.
# Reminder: Start time already declared at beginning of script.
# Now declaring and defining end time. And then subtracting to get timedelta.
def execution_time():
    print()
    print("Start Local Time: " + start_time.strftime("%A, %Y-%m-%d %H:%M:%S"))
    end_time = datetime.datetime.now()
    print("End Local Time: " + end_time.strftime("%A, %Y-%m-%d %H:%M:%S"))
    time_duration = end_time - start_time
    print("Total Time Elapsed in (Days) Hours:Mins:Secs: " + str(time_duration))

def end_results():
    print()
    # convert the list variables to string by type conversion to str
    # Strip the brackets from the lists by using .strip() method on the left and right bracket.

    # Only tell user to check empty Excel rows if any detected.
    if len(empty_rows) > 0:
        print("Check empty Excel rows: " + str(empty_rows).strip('[]'))
        print()
        
    # Only tell user to check missing entries rows if any detected.
    if len(missing_entries_rows) > 0:
        print("For copy to work, each row must be " \
              "filled out from source path to destination path.")
        print("Double check Excel rows: " + str(missing_entries_rows).strip('[]'))
        print()

    # Only tell user to check Excel rows with unmapped source paths if any detected.
    if len(unmapped_src_rows) > 0:
        print("Check Excel rows with unmapped source paths: " \
              + str(unmapped_src_rows).strip('[]'))
        print("To retain the desired folder structure, map" \
              "the directory above desired root folder as drive letter.")
        print()
        
    # Only tell user to Invalid paths rows if any detected.
    if len(invalid_paths_rows) > 0:
        print("Double check Excel rows with invalid paths: " \
              + str(invalid_paths_rows).strip('[]'))
        print('Make sure the source paths exist AND that each source path' \
              ' has a mapped drive letter above desired root folder.')
        print()

    # Always tell if there are any size mismatch errors resulting from copying
    print("Number of Size Mismatch Errors: " + str(size_mismatch_rows))

    # Only refer user to check errors.csv IF there are size mismatch errors.
    if size_mismatch_rows > 0:
        
        # sets the date and time in string format as a variable to later add to CSV name
        # This sets the time format in YearMonthDay_HrMinSec in Military Time
        now = time.strftime("%Y%m%d_%H%M%S")
        CSV_name = 'copy_errors_' + now + '.csv'
        
        # Generate the errors.csv
        # if you DO NOT set the newline parameter as '', then after the header, a row will be skipped.
            # This makes sure that you begin writing directly below the header.
        # It is good practice to use the with keyword when dealing with file objects.
            # The advantage is that the file is properly closed after its suite finishes, even if an exception is raised at some point.
        with open(CSV_name, 'w', newline='') as output_csv:
            fields = ['Source Path', 'Destination Path', 'Source Size', 'Destination Size']

            # The fieldnames parameter is a sequence of keys that identify the order in which
            # values in the dictionary passed to the writerow() method are written to the CSV file.
            # csv.DictWriter to set the header as a dictionary object that accepts values for its items
            Headers = csv.DictWriter(output_csv, fieldnames=fields)
            Headers.writeheader()

            # Make variable for writing to the CSV, so writing rows can begin
            output_writer = csv.writer(output_csv)

            # by now copy_errors should be a list of dictionary objects
                # each dictionary object with row header to values:
                # Ex. [{'Source Path': 'C:\\_Test\\new 2.txt', 'Destination Path': 'D:\\_TEST_TEST\\1\\', 'Source Size': 11, 'Destination Size': 23},
                # {'Source Path': 'C:\\_Test\\a\\', 'Destination Path': 'D:\\_TEST_TEST\\2', 'Source Size': 156, 'Destination Size': 149}]
            # .writerow() method normally accepts lists.
                # within each list element (the dictionary object), call on the value tied to the dictionary item.
            for row in copy_errors:
                output_writer.writerow([row['Source Path'], row['Destination Path'], \
                                        row['Source Size'], row['Destination Size']])

            print("See " + CSV_name + " if there were any size mismatch errors")

    print()

# Ask for user input so they can review results before closing.
    # Robocopy pauses if you right-click it. Left-click AND also the ENTER key will resume the copy.
    # However, pressing any button will cause the window to close if just asking for any user input.
    # Therefore, instead of asking simply ANY user input, ask for a specific input.
    # Only break while loop when specific input matched.
def user_close():
    while(True):
        user_input = input("To close this window, type 'exit' followed by ENTER" \
        " or click the close button: ")
        if user_input == 'exit':
            break

def main():
    variables()
    read_Excel()
    execution_time()
    end_results()
    user_close()

if __name__ == "__main__":
    try:
        main()
    except Exception:
        print()
        print("ERROR:")
        print(traceback.format_exc())
        input("Let Dev know of error. Screenshot error and keep Excel records.")
