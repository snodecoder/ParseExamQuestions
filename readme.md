## SYNOPSIS
Parse Word Document to JSON

## DESCRIPTION

Defines an Object structure from the predefined JSON structure created by Exam Simulator(https://github.com/exam-simulator).
Processes a Word Document paragraph by paragraph, strips unnecessary data, stores wanted data in objects, outputs JSON file.
Make sure that there is no variance in the source file in how for example a new question starts. If the script detects something it cannot process, it will output the last question  it was working on so you can verify and adjust the source file. When you've corrected the issue in the source file, you can then run the script again. If the script succesfully ran until the end, then you should have a clean output. 

## REQUIREMENTS
- Powershell 7 or later

*** TIP *** When you find errors in your output file or anthyng out of order that you would like to change, try doing so in the source file. Most of the time you will find another issue later on that requires you to make a adjustment to the source file and rerun it. By making all your afjustments in the source file, you prevent repeatedly changing your output file.

## EXAMPLE
How to use:

1. Edit the Default Parameters defined in script.
1. Structurize data in .docx file, so it has a consistent structure (example: Questions always start with 'Question [NUMBER]' / Answers start with 'A.' etc..).
1. Define the Selector Object below with text markers that identify the structure consistency.
1. Run Script

### Common Errors:
The script outputs last processed question so you can determine the cause of error.

Example: A new question part starts at the end of the previous paragraph (instead of starting on a new paragraphd)
  Bad:
    *"E.	Add Admin1 to the Enterprise Admins group. Correct Answer: B"*
  Good:
    *"E.	Add Admin1 to the Enterprise Admins group."*
    *"Correct Answer: B"*

## INPUTS
You input a .docx Word document.

## OUTPUTS
You get a.JSON file

## NOTES
To convert a PDF document to an editable word document, simply right click on the PDF file and choose Open With -> Word. Et voila!
No more hassle with mediocore online 'free' conversions, haha thanks MS!

Be aware that tweaking of the edit area in this script +   tweaking of the .docx file will be necessary to get a good final result. 

I mostly use the conversion to CSV nowadays, together with the free mobile app MTestM. The best working exam practice app I found so far (that enables you to bulk import questions)
