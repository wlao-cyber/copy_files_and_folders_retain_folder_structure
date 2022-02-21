
# Copy Files and Folders in Windows and Retain Their Folder Structure

This program reads off an Excel document, copying folders and files from user-specified source paths to user-specified destination paths while retaining the full folder structure or desired folder structure.

This will only work if the source path's root directory is a drive letter. To retain only a portion of the folder structure, you can also map a path, so that the desired folder shows as the top level folder. Network paths should be mapped.

For example: if you have <C:\Forensics\Test\Extracted_Files.zip\Exchange\Tbills.pst> and only want the top folder to be "Extracted_Files.zip", map the <C:\Forensics\Test\> folder as a drive letter. For more details on how to map network paths and local paths, see FEATURES, TIPS, and WARNINGS below.

Both local and long network paths should be accepted as the destination path, but the source path must contain a drive letter as the root directory.

If the destination path does NOT already exist, and the destination path is valid, the program will attempt to create the folder. Because of this feature, you can edit the destination folder path to include the desired folder structure.

The program harnesses Windows' built-in Robocopy tool to perform the copying. The following Robocopy settings were applied to retain all data, attributes, and time stamps: 
/COPY:DAT

## Authors

- [@wlao-cyber](https://github.com/wlao-cyber)


## FEATURES, TIPS, and WARNINGS:

- The program only accepts paths with drive letters. The console window will tell the user if no drive letter is detected.
- Network paths can be mapped from File Explorer by right-clicking "This PC" and selecting "Map network drive"
- Local paths can be mapped using the Windows subst command.

In Windows Command Prompt:
subst [unused drive letter]: [target path]
	
To unmap:
subst [mapped drive letter]: /D

- Before using the program again, completely delete all Excel rows below headers in case there were previous cells left with spaces or random characters.
    - Cells with just spaces (whitespace) will be reported as empty rows.
- Save Excel document with your changes prior to running EXE.
- Cells not detected as valid paths should be reported in results at the end of execution.
- The program autocorrects any folders/files that have leading or trailing whitespace, which are not allowed in Windows.
    - Folders cannot have leading or trailing whitespace
    - The file name should not have leading whitespace before the file name or trailing whitespace after the extension.
- Just like with Windows Command Prompt, if you left-click the console Window, it pauses Robocopy.
    - To resume, you can right-click or press ENTER.
- The program tries to identify invalid paths and will note them in the console window.
- A size mismatch errors CSV log will be generated if any size mismatches from copying are found.

## How to Use

- Download the EXE and XLSX pair of files locally and keep them in the same folder.
- Don't change the Excel XLSX Name. Otherwise, the program will not work.
- In the XLSX document, fill out the full source path and destination paths under the labelled columns
- Double click the EXE, and the program will start.
