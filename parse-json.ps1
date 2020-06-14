<#
.SYNOPSIS
  Parse Word Document to JSON
.DESCRIPTION
  Defines an Object structure from the predefined JSON structure (stored in .\functions.ps1).
  Processes a Word Document, strips unnecessary data, stores wanted data in objects, outputs JSON file.
.EXAMPLE
  PS C:\> <example usage>
  Explanation of what the example does
.INPUTS
  Inputs (if any)
.OUTPUTS
  Output (if any)
.NOTES
  General notes
#>

###############################
### >>> Start Edit Area >>> ###
###     Global Variables    ### 
###############################
param (
  $WordFileName = "742.docx"
  ,$folderPath = "C:\Codeprojects\ParseWordDocument\"
  ,$examNumber = "70-742"
  ,$imageURLPrefix = "https://files.doorhetgeluid.nl/images/$($examNumber)/"
)


$mediaFolder = "C:\Codeprojects\ParseWordDocument\$($WordFileName)\word\media\"
$imageFolder = $folderPath + "images\"
$reg = '([A-Z]{1})[\.](.*)' # Regex match string to select First letter in Option, replace '.' with ':)', finally add answer.
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
  )
} # End of Selector object

################################
### <<< End of Edit Area <<< ###
################################

### Load Modules ###
#. ($PSScriptRoot + ".\functions.ps1") # Load functions
$PSWriteWord = Get-InstalledModule -Name PSWriteWord -ErrorAction SilentlyContinue # Check if PSWriteWord is installed

if (!$PSWriteWord) {
  Install-Module -Name PSWriteWord -Force
}
Import-Module PSWriteWord -Force
Import-Module .\functions.psm1 -Force
### End Modules ###



$QuestionVariant = @{
  multipleChoice = 0 # [true,false,false,false]
  multipleAnswer = 1 # [true,true,false,false]
  fillInTheBlank = 2 # [answer,variation,another]
  listOrder = 3 # []
}
$NodeVariant = @{
  image = 0 # URL of an image
  text = 1 # Normal sized text, most commonly used variant
  largeText = 2 # Large header text
}

function insertVariant ($variant, $text) { # example use $jsonOutputObject.test[0].question += insertVariant $NodeVariant.text "dit is een test"
  New-Object [[PSCustomObject] -Property @{
    variant = $variant
    text = $text
  }]  
}

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

function insertChoice ($index, $text) { # example use $jsonOutputObject.test[0].question += insertVariant $NodeVariant.text "dit is een test"
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
  New-Object [PSCustomObject] -Property @{
    label = $label
    text = $text
  } 
}



######################## Process Word Document ########################
# Prepare Word Document for processing
$OldWordDocument = Get-WordDocument -FilePath ($folderPath + $WordFileName)
$paragraphs = $OldWordDocument.Paragraphs

# Create Image folder (for exported images) in working directory, if it not already exists
if ( (Test-Path -Path ($imageFolder)) -like "False" ) {
  New-Item -Path $folderPath -Name "images" -ItemType Directory
}
# Extract images from .docx file
extractWordImages -folderPath $folderPath -wordFileName $WordFileName


### Prepare Datastrucure ###
$QuestionArray = @()
$questid = 0
$QuestionArray += NewQuestion
$textExplanation = $false
# Prepare Exam
$jsonOutputObject = newJsonExam
$jsonOutputObject.test += NewJsonQuestion

# Store all the Question parts per Question in Objects, store Objects in $QuestionArray
for ( $i=0; $i -lt 23; $i++ ) {
  # write-host "starting round $($i)" # Turn on for Debugging

  if ( !($paragraphs[$i].text -like $Selector.question) ) { # If NOT start of new question, continue
    
    if ( ($paragraphs[$i].Pictures).count -like 1 ) # Images
    { 
      $jsonOutputObject.test[$questid].question += insertVariant -variant $NodeVariant.image -text $imageURLPrefix + $paragraphs[$i].Pictures.FileName
      
      $QuestionArray[$questid].image += $imageURLPrefix + $paragraphs[$i].Pictures.FileName
      Copy-Item -Path ($mediaFolder + $paragraphs[$i].Pictures.FileName) -Destination ($imageFolder + $paragraphs[$i].Pictures.FileName) -ErrorAction Ignore # Copy image to export folder for upload to server
    }
    elseif ($paragraphs[$i].text -like $Selector.filter ) # Filter unwanted text
    { 
      # skip it
    }  
    elseif ($paragraphs[$i].text -like $Selector.section ) # Section description of exam
    { 
        $QuestionArray[$questid].section += $paragraphs[$i].text
    } 
    elseif ( $paragraphs[$i].islistitem ) # Possible answers
    { 
      $choiceIndex = $jsonOutputObject.test[$questid].choices.Count
      $jsonOutputObject.test[$questid].choices += insertChoice -index $choiceIndex -text $paragraphs[$i].text
      
      $QuestionArray[$questid].answers += $paragraphs[$i].text
    }
    elseif ( $paragraphs[$i].text -like $Selector.correct ) # Correct answer
    {
      $CorrectAnswer = ($paragraphs[$i].text).replace("Correct Answer: ","")
      $QuestionArray[$questid].correct += $CorrectAnswer.ToCharArray()
      
    }
    elseif ( $textExplanation ) # Add to Explanation Array
    {
        $QuestionArray[$questid].explanation += $paragraphs[$i].text
    }
    elseif ($paragraphs[$i].text -like $Selector.explanation ) # Add to explanation property
    {
        $QuestionArray[$questid].explanation += $paragraphs[$i].text
        $textExplanation = $true # Ensures all in-question-buffer is stored in Explanation array.
        $jsonOutputObject.test[$questid].explanation += $paragraphs[$i].text
    }
    else # The question itself
    { 
      $QuestionArray[$questid].text += $paragraphs[$i].text
      $jsonOutputObject.test[$questid].question += insertVariant -variant $NodeVariant.text -text $paragraphs[$i].text
    }
  }
  elseif ( (Like $paragraphs[$i].text $Selector.question) ) { # New question starts, reset everything

    if ($QuestionArray[$questid].correct[0].Length -like 1){
        $QuestionArray[$questid].type = "single_answer"
        $jsonOutputObject.test[$questid].variant = 0
    }
    elseif ($QuestionArray[$questid].correct[0].Length -gt 1){
        $QuestionArray[$questid].type = "multiple_answers"
        $jsonOutputObject.test[$questid].variant = 1
    }

    $jsonOutputObject.test[$questid].answer = booleanAnswer -CorrectAnswers $QuestionArray[$questid].correct -ChoicesCount $QuestionArray[$questid].answers.Count # Save array of answers in boolean
    $QuestionArray[$questid].index = $questid
    $QuestionArray += NewQuestion
    $jsonOutputObject.test += NewJsonQuestion
    $textExplanation = $false
    $questid ++
  } 
} # End for loop


$jsonOutputObject

$OutputObject = @()
for ($i=1; $i -lt $jsonOutputObject.test.Count; $i++) {
  $OutputObject += $jsonOutputObject.test[$i]
}

$jsonExport = $jsonOutputObject | ConvertTo-Json -Depth 6 -AsArray -EnumsAsStrings

[array[]]$jsonQuestions = @()

$jsonQuestions += $OutputObject[1] | ConvertTo-Json -Depth 3 -AsArray -EnumsAsStrings
$jsonExport = $jsonQuestions 
$jsonExport | Out-File -FilePath ($folderPath + "new-$($examNumber).json") -Force




  # Save Data as JSON
  #$QuestionArray | ConvertTo-Json | Out-File -FilePath ($folderPath + "new-$($examNumber).json")





