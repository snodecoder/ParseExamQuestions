##.SYNOPSIS
Parse Word Document to JSON

##.DESCRIPTION
Defines an Object structure from the predefined JSON structure created by Exam Simulator(https://github.com/exam-simulator).
Processes a Word Document paragraph by paragraph, strips unnecessary data, stores wanted data in objects, outputs JSON file.

##.EXAMPLE
How to use:

1. Edit the Default Parameters defined in script.
1. Structurize data in .docx file, so it has a consistent structure (example: Questions always start with 'Question [NUMBER]' / Answers start with 'A.' etc..).
1. Define the Selector Object below with text markers that identify the structure consistency.
1. Run Script

###Common Errors:
The script outputs last processed question so you can determine the cause of error.

Example: A new question part starts at the end of the previous paragraph (instead of starting on a new paragraphd)
  Bad:
    *"E.	Add Admin1 to the Enterprise Admins group. Correct Answer: B"*
  Good:
    *"E.	Add Admin1 to the Enterprise Admins group."*
    *"Correct Answer: B"*

##.INPUTS
You input a .docx Word document.

##.OUTPUTS
You get a.JSON file

##.NOTES
I've had success converting several PDF exams to .docx (with https://www.sejda.com/), and then finally to .JSON. Be aware that tweaking of the edit area in this script +   tweaking of the .docx file will be necessary to get a good final result. 
#>
