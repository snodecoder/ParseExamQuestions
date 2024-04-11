## SYNOPSIS
Parse Exam Questions Examtopics

## DESCRIPTION

Crawls Examtopics.com for all the relevant exam questions (search is based on the provided $ExamCode).
The questions will be exported in CSV format, compatible with the exam testing app MTestM (Android & iOS)

## REQUIREMENTS
- Selenium Chrome Webdriver (same version as your Google Chrome browser) https://googlechromelabs.github.io/chrome-for-testing/
- Powershell 7 or later


*** TIP *** When you find errors in your output file or anything out of order that you would like to change, try doing so by tweaking the code instead of the CSV. Most of the time you will find another issue later on that requires you to make a adjustment to the source file and rerun it. By making all your adjustments in the source file, you prevent repeatedly changing your output file.

## EXAMPLE
How to use:

1. Update Selenium chromedriver.exe so that it's version corresponds with your installed Google Chrome browser's version
1. Run Script

### Common Errors:
Sometimes the location of specific items on the website changes, which leads to a broken script. If this happens use DevTools in your browser on the website to update the corresponding FullXPath of the item that caused the error.

## INPUTS
ExamCode used as search parameter for crawling exam questions

## OUTPUTS
A CSV file formatted for use with MTestM app for mobile devices (Android & iOS)

## Issues / Contributing
Please log any issues you find and feel free to contribute :)

