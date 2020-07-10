<#
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

###############################
### >>> Start Edit Area >>> ###
###    Default Parameters   ###
###############################
param (
  $examAuthorId = "00001"
  ,$examAuthorName = "Snodecoder"
  ,$examAuthorImage = "http://www.example.com/image.png"
  ,$examCode= "70-742"
  ,$examTitle = "Identity with Windows Server 2016"
  ,$examDescription = "161 questions available in Multiple Choice en Multiple Answer format."
  ,$examImage = "http://www.example.com/image.png"
  ,$examTime = 120 # Maximum time for exam
  ,$examPass = 75 # Minimum percentage to pass exam
  ,$imageURLPrefix = "https://start.opensourceexams.org/exams/$($examCode)/images/"
  ,$WordFileName = "742.docx"
  ,$folderPath = "C:\Codeprojects\ParseWordDocument\"
)

try{
  ##################################
  ### >>> Continue Edit Area >>> ###
  ###      Global Variables      ###
  ##################################
  $Selector = New-Object psobject -Property @{
    question = "QUESTION*"
    ;explanation = "Explanation*"
    ;correct = "Correct Answer*"
    ;section = "Section*"
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
      ,"*Note: This question*"
      ,"*Start of repeated scenario*"
      ,"*End of repeated scenario*"
      ,"*After you answer a question in this section*"
    )
  } # End of Selector object

  ################################
  ### <<< End of Edit Area <<< ###
  ################################
  $mediaFolder = $folderPath + "$($WordFileName.Remove(3,5))\word\media\"
  $imageFolder = $folderPath + "images\"
  $WarningPreference = 'Continue'



  ### Load Modules ###
  #. ($PSScriptRoot + ".\functions.ps1") # Load functions
  $PSWriteWord = Get-InstalledModule -Name PSWriteWord -ErrorAction SilentlyContinue # Check if PSWriteWord is installed

  if (!$PSWriteWord) {
    Install-Module -Name PSWriteWord -Force
  }
  Import-Module PSWriteWord -Force

  ### End Modules ###

  ##### Functions & Class DEFINITIONS #####
  class TextVariant # Text Vvariant (Large, Normal, Url)
  {
    [int] $variant
    [string] $text

    TextVariant([int] $variant, [string] $text)
    {
      $this.variant = $variant
      $this.text = $text
    }
  } # End class TextVariant

  class TextLabel # Text Label for choices (A, B, C...)
  {
    [string] $label
    [string] $text

    TextLabel([string] $label, [string] $text)
    {
      $this.label = $label
      $this.text = $text
    }
  } # End class TextLabel

  class Question # Question constructor
  {
    [int] $variant
    [array] $question
    [array] $choices
    [array] $answer
    [array] $explanation

    Question() # Constructor
    {
      $this.variant # question variant
      $this.question # body of actual question
      $this.choices # body of actual choices
      $this.answer # array with true/false for every choice
      $this.explanation # explanation
    }
  }

  function NewJsonExam () {
    [PSCustomObject]@{
      title = [string]$null # exam title
      description = [string]$null # exam description
      author = [PSCustomObject]@{
        id = [string]$null # author ID
        name = [string]$null # author name
        image = [string]$null # author image
      }
      createdAt = [datetime] # creation datetime
      code = [string]$null # exam number
      time = [int]$null # maximum exam time
      pass = [int]$null # minimum score required to pass exam
      image = [string]$null # cover image of exam
      cover = [array[]] @() # fill array with addText method
      test = [array[]] @() # stores questions via addQuestion method
    }
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

  function ConvertAnswer($answer) {
    $input = $answer.tostring()
    switch ( $input ) {
      "A" {"0"; Break}
      "B" {"1"; break}
      "C" {"2"; break}
      "D" {"3"; break}
      "E" {"4"; break}
      "F" {"5"; break}
      "G" {"6"; break}
      "H" {"7"; break}
      "I" {"8"; break}
      "J" {"9"; break}
      "K" {"10"; break}
      "L" {"11"; break}
    }
  } # End of function ConvertAnswer

  function booleanAnswer ($CorrectAnswers, $ChoicesCount) { # Generate Array with true or false (if correct answer) for each answer
    [System.Boolean[]]$booleanAnswers = @()
    [int[]]$correct = @()

    $CorrectAnswers | ForEach-Object { # convert Correct Character answer (A, or B) to decimal index
      $correct += ConvertAnswer $_
    }

    for ($i = 0; $i -lt $ChoicesCount; $i++) { # generate true if decimal index correct == index of choices, otherwise false
      $booleanAnswers += $correct.Contains($i)
    }
    $booleanAnswers
  }

  function AddChoice ($index, $text) { # example use $jsonOutputObject.test[0].question += insertVariant $NodeVariant.text "dit is een test"
    switch ($index) {
      0 {$label = "A"}
      1 {$label = "B"}
      2 {$label = "C"}
      3 {$label = "D"}
      4 {$label = "E"}
      5 {$label = "F"}
      6 {$label = "G"}
      7 {$label = "H"}
      8 {$label = "I"}
      9 {$label = "J"}
      10 {$label = "K"}
      11 {$label = "L"}
      Default {}
    }
    [TextLabel]::new($label, $text)
  }

  function AddTextVariant () { # Helper function to add textVariant blocks
    param(
      [Parameter(Mandatory=$true,
      HelpMessage="0=Image URL, 1=Normal Size, 2=Large Size")]
      [ValidateSet("ImageURL" , "Normal", "Large")]
      [string]$variant,
      [Parameter(Mandatory=$false,
      HelpMessage="Enter Text")]
      [string]$text
    )

    [int]$intVariant = switch ($variant) { # Convert to decimal
      ImageURL { 0 }
      Normal { 1 }
      Large { 2 }
      Default {}
    }

    if ($text.Length -like 0) { # add space to be able to store a blank line of text
      $text = " "
    }

    [TextVariant]::new($intVariant, $text)
  }

  function addQuestionType () { # Helper function to add QuestionType
    param(
      [Parameter(Mandatory=$true,
      HelpMessage="Choose type of question")]
      [ValidateSet("MultipleChoice", "MultipleAnswer", "FillInTheBlank", "ListOrder")]
      [string]$type
    )
    [int]$intType = switch ($type) {
      MultipleChoice { 0 }
      MultipleAnswer { 1 }
      FillInTheBlank { 2 }
      ListOrder { 3 }
      Default {}
    }
    $intType
  }

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


  ########## Prepare Datastrucure ##########
  $questid = 0
  $textExplanation = $false
  $exam = newJsonExam
  $exam.title = $examTitle
  $exam.description = $examDescription
  $exam.author.id = $examAuthorId
  $exam.author.name = $examAuthorName
  $exam.author.image = $examAuthorImage
  $exam.createdAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
  $exam.image = $examImage
  $exam.code = $examCode
  $exam.time = $examTime
  $exam.pass = $examPass
  $exam.cover += AddTextVariant -variant Large -text $examTitle
  $exam.cover += AddTextVariant -variant Normal -text $examDescription
  $exam.test += [Question]::new()


########## PROCESS TEXT ##########
  # Store all the Question parts per Question in Objects, store Objects in $QuestionArray
  for ( $i=0; $i -lt $paragraphs.Count; $i++ ) {
    # write-host "starting round $($i)" # Turn on for Debugging

    if ( !($paragraphs[$i].text -like $Selector.question) ) { # If NOT start of new question, continue

##### IMAGES #####
      if ( ($paragraphs[$i].Pictures).count -like 1 ) {
        if ( $paragraphs[$i].Pictures.width -lt 15 ) {
          # Skip it, non-relevant image
        }
        else {  # Store imagelink in text for current question. Copy image file to export folder, upload this to tje $imageURLPrefix location on your webserver
          $exam.test[$questid].question += AddTextVariant -variant ImageURL -text ($imageURLPrefix + $paragraphs[$i].Pictures.FileName)
          Copy-Item -Path ($mediaFolder + $paragraphs[$i].Pictures.FileName) -Destination ($imageFolder + $paragraphs[$i].Pictures.FileName) -ErrorAction Ignore # Copy image to export folder for upload to server
        }
      }

##### FILTER #####
      elseif ( Like -string $paragraphs[$i].text -arrayStrings $Selector.filter ) { # Uses Like Function to search if current text exists in array of text. -> Filter unwanted text
        # skip it
      }

##### SECTION #####
      elseif ( $paragraphs[$i].text -like $Selector.section ) { # Section description of exam
        # skip it
      }

##### CORRECT ANSWERs #####
      elseif ( $paragraphs[$i].text -like $Selector.correct ) { # Correct answer: Convert correct answers to boolean array and store in $exam
        $CorrectAnswer = ($paragraphs[$i].text).replace("Correct Answer: ","")
        $exam.test[$questid].answer = booleanAnswer -CorrectAnswers ($CorrectAnswer.ToCharArray()) -ChoicesCount ($exam.test[$questid].choices.count)

        # Determine type of question
        if ( $CorrectAnswer.Length -like 1 ) {
          $exam.test[$questid].variant = addQuestionType -type MultipleChoice
        }
        elseif ( $CorrectAnswer.Length -gt 1) {
          $exam.test[$questid].variant = addQuestionType -type MultipleAnswer
        }
      }

##### EXPLANATION #####
      elseif ( $textExplanation ) { # Add to Explanation Array

        if ( $paragraphs[$i].text -like "*http*") { # Place HTTP/HTTPS links on new line
          $subStrings = $paragraphs[$i].Text.Split(" ")
          foreach ($string in $subStrings) {
            $exam.test[$questid].explanation += AddTextVariant -variant Normal -text $string
          }
        }
      }
      elseif ( $paragraphs[$i].text -like $Selector.explanation ) { # Add to explanation property
          $textExplanation = $true # Ensures all in-question-buffer is stored in Explanation array.
          $exam.test[$questid].explanation += AddTextVariant -variant Normal -text $paragraphs[$i].text
      }

##### POSSIBLE ANSWERS #####
      elseif ( ($paragraphs[$i].islistitem) -or (Like -string $paragraphs[$i].text -arrayStrings $Selector.options) ) { # Possible answers: Store available choices in question
        $choiceIndex = $exam.test[$questid].choices.Count
        $exam.test[$questid].choices += AddChoice -index $choiceIndex -text $paragraphs[$i].text
      }

##### ACTUAL QUESTION #####
      else { # The question itself
        $exam.test[$questid].question += AddTextVariant -variant Normal -text $paragraphs[$i].text
      }
    }

##### NEW QUESTION: RESET #####
    elseif ( (Like $paragraphs[$i].text $Selector.question) ) { # New question starts, reset everything++

      if ( $exam.test[$questid].question -like $null ) { # if question was not filled, recycle it
        $exam.test[$questid] = [Question]::new()
      }
      elseif ($exam.test[$questid].answer -notcontains "True") {
        throw "Correct answer array does not contain a correct answer for question ($($questid + 1)) in .docx file."
      }
      else {
        $exam.test += [Question]::new() # add new empty Question object to exam
        $questid ++ # increment Question ID for processing next question
      }
      $textExplanation = $false # reset the textexplanation value
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




########## Convert Exam to JSON and Export it ##########
$jsonExam = $exam | ConvertTo-Json -Depth 5 -Verbose

if ( $jsonExam | Test-Json ) {
  $jsonExam | Out-File -FilePath ($folderPath + "$($examCode).json") -Force
  Write-Host "Exported $($exam.test.Count) questions to JSON file :)" -ForegroundColor Green
}
else {
  Write-Warning "Please check generated jsonExam. It is not a valid JSON file."
}
