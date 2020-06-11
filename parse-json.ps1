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
  $OldWord = "742.docx"
  ,$folderPath = "C:\Codeprojects\ParseWordDocument\"
  ,$examNumber = "70-742"
  ,$imageURLPrefix = "https://files.doorhetgeluid.nl/images/$($examNumber)/"
)


$mediaFolder = "C:\Codeprojects\ParseWordDocument\docx\word\media\"
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
. ($PSScriptRoot + ".\functions.ps1") # Load functions
$PSWriteWord = Get-InstalledModule -Name PSWriteWord -ErrorAction SilentlyContinue # Check if PSWriteWord is installed

if (!$PSWriteWord) {
  Install-Module -Name PSWriteWord -Force
}
Import-Module PSWriteWord -Force 
### End Modules ###

### Prepare Datastrucure ###
$jsonOutputObject = newJsonExam
$jsonOutputObject.author = ""
$QuestionArray = @()
$questid = 0
$QuestionArray += NewQuestion
$textExplanation = $false


######################## Process Word Document ########################
# Prepare Word Document for processing
$OldWordDocument = Get-WordDocument -FilePath ($folderPath + $OldWord)
$paragraphs = $OldWordDocument.Paragraphs

# Create Image folder (for exported images) in working directory, if it not already exists
 if ( (Test-Path -Path ($imageFolder)) -like "False" ) {
   New-Item -Path $folderPath -Name "images" -ItemType Directory
 }
extractWordImages -folderPath $folderPath -wordFileName $OldWord


# Store all the Question parts per Question in Objects, store Objects in $QuestionArray
for ( $i=0; $i -lt $paragraphs.count; $i++ ) {
  # write-host "starting round $($i)" # Turn on for Debugging

  if ( !($paragraphs[$i].text -like $Selector.question) ) { # If NOT start of new question, continue
    
    if ( ($paragraphs[$i].Pictures).count -like 1 ) # Images
    { 
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
        $QuestionArray[$questid].answers += $paragraphs[$i].text
    }
    elseif ( $paragraphs[$i].text -like $Selector.correct ) # Correct answer
    {
        $QuestionArray[$questid].correct += ($paragraphs[$i].text).replace("Correct Answer: ","")
    }
    elseif ( $textExplanation ) # Add to Explanation Array
    {
        $QuestionArray[$questid].explanation += $paragraphs[$i].text
    }
    elseif ($paragraphs[$i].text -like $Selector.explanation ) # Add to explanation property
    {
        $QuestionArray[$questid].explanation += $paragraphs[$i].text
        $textExplanation = $true # Ensures all in-question-buffer is stored in Explanation array.
    }
    else # The question itself
    { 
      $QuestionArray[$questid].text += $paragraphs[$i].text
    }
  }
  elseif ( (Like $paragraphs[$i].text $Selector.question) ) { # New question starts, reset everything

    if ($QuestionArray[$questid].correct[0].Length -like 1){
        $QuestionArray[$questid].type = "single_answer"
    }
    elseif ($QuestionArray[$questid].correct[0].Length -gt 1){
        $QuestionArray[$questid].type = "multiple_answers"
    }

    $QuestionArray[$questid].index = $questid
    $QuestionArray += NewQuestion
    $textExplanation = $false
    $questid ++
  } 
} # End for loop

$QuestionArray




  # Save Data as JSON
  $QuestionArray | ConvertTo-Json | Out-File -FilePath ($folderPath + "new-$($examNumber).json")





