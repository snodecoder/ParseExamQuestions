#
.SYNOPSIS
  Parse Word Document to JSON
.DESCRIPTION
  Defines an Object structure from the predefined JSON structure created by Exam Simulator(https://github.com/exam-simulator).
  Processes a Word Document paragraph by paragraph, strips unnecessary data, stores wanted data in objects, outputs JSON file.
.EXAMPLE
  How to use:

  1. Edit the Default Parameters defined in script.
  2. Structurize data in .docx file, so it has a consistent structure (example: Questions always start with 'Question [NUMBER]' / Answers start with 'A.' etc..).
  3. Define the Selector Object below with text markers that identify the structure consistency.
  4. Run Script

  Common Errors:
  The script outputs last processed question so you can determine the cause of error.

  Example: A new question part starts at the end of the previous paragraph (instead of starting on a new paragraphd)
    Bad:
      "E.	Add Admin1 to the Enterprise Admins group. Correct Answer: B"
    Good:
      "E.	Add Admin1 to the Enterprise Admins group."
      "Correct Answer: B"


.INPUTS
  You input a .docx Word document
.OUTPUTS
  .JSON file
.NOTES
  General notes
#>