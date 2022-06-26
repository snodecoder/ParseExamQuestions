<#
.SYNOPSIS
  Parse Word Document to CSV
.DESCRIPTION
  Processes a Word Document paragraph by paragraph, strips unnecessary data, stores wanted data in objects, outputs CSV file in MTestM format (app is available from Playstore)
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
  .CSV File
.NOTES
  General notes
#>

###############################
### >>> Start Edit Area >>> ###
###    Default Parameters   ###
###############################
# Adds the general information for your practice exam.

param (
  $examCode= "AZ104"
  ,$examTitle = "AZ104"
  ,$examDescription = "Practice questions in Multiple Choice en Multiple Answer format."
  ,$examDuration = 120 # Maximum time for exam
  ,$examKeywords = "Azure, Fundamentals"
  ,$imageURLPrefix = "https://start.opensourceexams.org/exams/$($examCode)/images/"
  ,$WordFileName = "$($examCode).docx"
  ,$folderPath = "C:\CodeProjects\ParseWordDocument\"
)



try{
  ##################################
  ### >>> Continue Edit Area >>> ###
  ###      Global Variables      ###
  ##################################
  # This is where you tell the script how the recognize the start of a new question, explanation answer, etc..
  # Make sure that there is no variance in the source file in how for example a new question starts. If the script detects something it cannot process, it will output the last question  it was working on so you can verify and adjust the source file. When you've corrected the issue in the source file, you can then run the script again. If the script succesfully ran until the end, then you should have a clean output.

  # *** TIP *** When you find errors in your output file or anthyng out of order that you would like to change, try doing so in the source file. Most of the time you will find another issue later on that requires you to make a adjustment to the source file and rerun it. By making all your afjustments in the source file, you prevent repeatedly changing your output file.

  $Selector = New-Object psobject -Property @{
    question = "New Question *"
    ;explanation = "Explanation:*"
    ;correct = " Answer:*"
    ;section = "- (Exam Topic*"
    ;options = @(
      "A.*"
      ,"B.*"
      ,"C.*"
      ,"D.*"
      ,"E.*"
      ,"F.*"
      ,"G.*"
      ,"H.*"
    )
    ;imageFormat = @(
      "*.jpeg"
      ,"*.png"
    )
    ;filter = @(
      "*gratisexam*"
      ,"*topic*"
      ,"*Note: This question*"
      ,"*Start of repeated scenario*"
      ,"*End of repeated scenario*"
      ,"*After you answer a question in this section*"
    )
    ;type = @(
      "*hotspot*"
      ,"*drag drop*"
    )
  } # End of Selector object

  ################################
  ### <<< End of Edit Area <<< ###
  ################################
  $mediaFolder = $folderPath + "$($WordFileName.TrimEnd(".docx"))\word\media\"
  $imageFolder = $folderPath + "images\"
  $WarningPreference = 'Continue'
  $ActiveSection = ""



  ### Load Modules ###
  #. ($PSScriptRoot + ".\functions.ps1") # Load functions
  $PSWriteWord = Get-InstalledModule -Name PSWriteWord -ErrorAction SilentlyContinue # Check if PSWriteWord is installed

  if (!$PSWriteWord) {
    Install-Module -Name PSWriteWord -Force
  }
  Import-Module PSWriteWord -Force

  ### End Modules ###

  ##### Functions & Class DEFINITIONS #####


  class Question # Question constructor
  {
    [string]$section
    [string]$material
    [string]$question
    [string]$no
    [string]$type
    [string]$option1
    [string]$option2
    [string]$option3
    [string]$option4
    [string]$option5
    [string]$option6
    [string]$option7
    #[string]$option8
    #[string]$option9
    #[string]$option10
    #[string]$option11
    #[string]$option12
    [string]$explanation
    [string]$answer1
    [string]$answer2
    [string]$answer3
    [string]$answer4
    [string]$score1
    [string]$score2
    [string]$score3
    [string]$score4

    Question() # Constructor
    {
      $this.section
      $this.material
      $this.question
      $this.no
      $this.type
      $this.option1
      $this.option2
      $this.option3
      $this.option4
      $this.option5
      $this.option6
      $this.option7
      $this.explanation
      $this.answer1
      $this.answer2
      $this.answer3
      $this.answer4
      $this.score1
      $this.score2
      $this.score3
      $this.score4
    }
  }

  function NewCSVExam () {
    [array[]]@()
  } # End of function newJsonExam

  function Like ( $string, $arrayStrings ) { # Perform like search in Array
    $result = $false

    $arrayStrings | ForEach-Object {
      if ($string -ilike $_ ) {
        $result = $true
      }
    }
    return $result
  } # End of function Like

  function ExtractWordImages($folderPath, $wordFileName) { # extracts images from .docx and stores them in .\images folder,
    $wordFile = Get-ChildItem -Path ($folderPath + $wordFileName) -Filter *.docx
    Rename-Item $wordFile -NewName ($wordFile.BaseName + ".zip")
    Expand-Archive ($wordFile.BaseName + ".zip") -Force

    $zipFile = Get-ChildItem -Path ($folderPath + $wordFile.BaseName + ".zip") -Filter *.zip
    Rename-Item -Path $zipFile.FullName -NewName ($zipFile.BaseName + ".docx")
  } # End of function extractWordImages
}
catch{
  Write-Warning -Message "$($_) : Error in setting up Global Variables, Modules, Classed and Functions. Please review."
}


try {
  ########## Process Word Document ##########
  # Prepare Word Document for processing
  $OldWordDocument = Get-WordDocument -FilePath ($folderPath + $WordFileName)
  $paragraphs = $OldWordDocument.Paragraphs

  # Create Image folder (for exported images) in working directory, if it not already exists
  if ( (Test-Path -Path ($imageFolder)) -like "False" ) {
    New-Item -Path $folderPath -Name "images" -ItemType Directory | Out-Null
  }
  elseif ( (Test-Path -Path ($imageFolder)) -like "True" ) {
    Remove-Item -Path $imageFolder -Recurse
    New-Item -Path $folderPath -Name "images" -ItemType Directory | Out-Null
  }
  # Extract images from .docx file
  extractWordImages -folderPath $folderPath -wordFileName $WordFileName



# Prepare data structure
$questid = 0
$textExplanation = $false
[array]$exam = @()
$QuestionObject = [Question]::new()
$index = 0
$tempOptions = $null

########## PROCESS TEXT ##########
  # Store all the Question parts per Question in Objects, store Objects in $QuestionArray
  for ( $i=0; $i -lt $paragraphs.Count; $i++ ) {

    #write-host "Processing paragraph: $($i)." # Turn on for Debugging

    if ( !($paragraphs[$i].text -like $Selector.question) ) { # If NOT start of new question, continue

##### IMAGES #####
      if ( ($paragraphs[$i].Pictures).count -like 1 ) {
        if ( $paragraphs[$i].Pictures.width -lt 15 ) {
          # Skip it, non-relevant image
        }
        else {  # Store imagelink in text for current question. Copy image file to export folder, upload this to tje $imageURLPrefix location on your webserver
        $i
          $QuestionObject.question += "<img src='$($imageURLPrefix + $paragraphs[$i].Pictures.FileName)' style='max-width: 100%;'>"
          Copy-Item -Path ($mediaFolder + $paragraphs[$i].Pictures.FileName) -Destination ($imageFolder + $paragraphs[$i].Pictures.FileName) -ErrorAction Ignore # Copy image to export folder for upload to server
        }
      }

##### FILTER #####
      elseif ( Like -string $paragraphs[$i].text -arrayStrings $Selector.filter ) { # Uses Like Function to search if current text exists in array of text. -> Filter unwanted text
        # skip it
      }

      elseif ( $paragraphs[$i].text.Length -like 0 ) {
        # skip it
      }

##### SECTION #####
      elseif ( $paragraphs[$i].text -like $Selector.section ) { # Section description of exam
        $QuestionObject.section = $paragraphs[$i].text
      }

##### CORRECT ANSWERs #####
      elseif ( $paragraphs[$i].text -like $Selector.correct ) { # Correct answer: Convert correct answers to boolean array and store in $exam
        $CorrectAnswer = ($paragraphs[$i].text).replace("Correct Answer: ","")
        $QuestionObject.answer1 = $CorrectAnswer

        # Determine type of question
        if ( $CorrectAnswer.Length -like 1 ) {
          $QuestionObject.type = ""
        }
        elseif ( $CorrectAnswer.Length -gt 1) {
          $QuestionObject.type = "multiple choice" # When multiple answers must be given
        }
        $textExplanation = $true
      }

##### EXPLANATION #####
      elseif ( $textExplanation ) { # Add to Explanation Array

        if ( $paragraphs[$i].Text -like "*https://books.google.co.*") {
          $textExplanation = $false
        }
        elseif ( $paragraphs[$i].text -like "*http*") { # Place HTTP/HTTPS links on new line
          $subStrings = $paragraphs[$i].Text.Split(" ")
          foreach ($string in $subStrings) {
            $QuestionObject.explanation += "<a href='$($string)'>Link</a>"
          }
        }
        elseif ( $paragraphs[$i].text.Length -gt 0 ) {
          $QuestionObject.explanation += "<p>$($paragraphs[$i].text)</p>"
        }
      }
      elseif ( $paragraphs[$i].text -like $Selector.explanation ) { # Add to explanation property
          $textExplanation = $true # Ensures all in-question-buffer is stored in Explanation array.
          $QuestionObject.explanation += "<p>$($paragraphs[$i].text)</p>"
      }

##### POSSIBLE ANSWERS #####
      elseif ( ($paragraphs[$i].islistitem) -or (Like -string $paragraphs[$i].text -arrayStrings $Selector.options) ) { # Possible answers: Store available choices in question
        $tempOptions += $paragraphs[$i].text
      }

##### ACTUAL QUESTION #####
      else { # The question itself
        $QuestionObject.question += "<p>$($paragraphs[$i].text)</p>"
      }
    }

##### NEW QUESTION: RESET #####
    elseif ( (Like $paragraphs[$i].text $Selector.question) ) { # New question starts, reset everything++

      for ($ii = 0; $ii -lt $tempOptions.Count; $ii++) {
        $item = "option$($ii+1)"
        $QuestionObject.$item = $tempOptions[$ii]
      }
      if ( ! (Like $QuestionObject.question $Selector.type) ) { # do not export hot spot / drag drop questions
      }

      # add extra points when multiple answers are corect
      $numberAnswers = $QuestionObject.answer1.Length
      for ($ii = 1; $ii -le $numberAnswers; $ii++) {
        $item = "score$($ii)"
        $QuestionObject.$item = "1.0"
      }

      if ($ActiveSection -notlike $QuestionObject.section) # delete section title if section is same as section previous question
      {
        $section = [Question]::new()
        $section.section = $QuestionObject.section
        $exam += $section
        $ActiveSection = $QuestionObject.section
        $QuestionObject.section = ""
      }
      elseif ($ActiveSection -like $QuestionObject.section)
      {
        $QuestionObject.section = ""
      }

      if ($QuestionObject.question.length -gt 0) { # store questionobject in Exam object
        $exam += $QuestionObject
      }



      # Reset for next question
      $QuestionObject = [Question]::new()
      $textExplanation = $false # reset the textexplanation value
      $tempOptions = @()
      $index++
    }
##### START LOOP AGAIN #####

  } # End for loop
  Write-Host "Finished processing document." -ForegroundColor Green
}

catch{
  Write-Warning -Message "$($_): in executing Paragraph: $($i) conversion."

  Write-Host "Question summary: " -ForegroundColor Blue
  $exam.test[$questid]
  Write-Host "Question: Text" -ForegroundColor Blue
  $exam.test[$questid].question | Format-Table
  Write-Host "Question: Choices" -ForegroundColor Blue
  $exam.test[$questid].choices | Format-Table
  Write-Host "Question: Answers" -ForegroundColor Blue
  $exam.test[$questid].answer | Format-List
  Write-Host "Question: Explanation" -ForegroundColor Blue
  $exam.test[$questid].explanation

  Write-Host "Please fix consistency problem in structure .docx file." -ForegroundColor Red
  Write-Host "Mostly this is caused by a question part (for example 'Correct Answer: A') that does not have it's own 'line', fix by placing question part on a new line with enter."
  exit 0
}
########## FINISHED PROCESSING DOCUMENT ##########


[array]$Output = @()

foreach ($item in $exam) {

  if ($item.question.length -gt 0) {
    $Output += $item
  }

}

function printSemiColon ($columnCount)
{
  [string]$output = ""

  for ($i=0; $i -lt ($columnCount -2); $i++)
  {
    $output += ";"
  }
  return $output
}

########## Convert Exam to CSV and Export it ##########
$ColumnCount = ($exam | Get-Member -MemberType Property).count

$header = "Title;$($examTitle)$(printSemiColon $ColumnCount `n)`
Description;$($examDescription)$(printSemiColon $ColumnCount `n)`
Duration;$($examDuration)$(printSemiColon $ColumnCount `n)`
Keywords;$($examKeywords)$(printSemiColon $ColumnCount `n)"

$header | Out-File -FilePath ($folderPath + "$($examCode).csv") -Force

$CSV = $exam | ConvertTo-Csv -Delimiter ";" -UseQuotes Never
$CSV | Out-File -FilePath ($folderPath + "$($examCode).csv") -Append

#$CSV | Export-Csv -path ($folderPath + "$($examCode).CSV") -Delimiter ";" -UseQuotes Never -Encoding unicode


#  $CSV | Out-File -FilePath ($folderPath + "$($examCode).CSV") -Force
  Write-Host "Exported $($exam.Count) questions to CSV file :)" -ForegroundColor Green


