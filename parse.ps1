$OldWord = "test.docx"
$folderPath = "C:\Codeprojects\ParseWordDocument\"
$imageFolder = "C:\Codeprojects\ParseWordDocument\test\word\media\"

$reg = '([A-Z]{1})[\.](.*)' # Regex match string to select First letter in Option, replace '.' with ':)', finally add answer.
$Selector = New-Object psobject -Property @{
  question = "QUESTION*"
  ;explanation = "Explanation*"
  ;correct = "Correct Answer*"
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

# install PSWriteWord module to easier edit word document: "install-module -name PSWriteWord -Force"
Import-Module PSWriteWord -Force

# Functions
function NewQuestion (){ # Create new question object
  $propertylist = @{
    text = @()
    ;image = @()
    ;answers = @()
    ;correct = @()
    ;explanation = @()
  }
  $question = New-Object psobject -Property $propertylist
  $question
} # End of function NewQuestion

function Like ( $str, $patterns ) { # Perform like search in Array
  $patterns | ForEach-Object {
    if ($str -ilike $_ ) {
      return $true
    }
  }  
} # End of function Like

# Prepare Word Document
$WordDocument = New-WordDocument ($folderPath + "new.docx")
$OldWordDocument = Get-WordDocument -FilePath ($folderPath + $OldWord)
$paragraphs = $OldWordDocument.Paragraphs

# Prepare data structure
$buffer = @{}
$QuestionArray = @()
$questionIndex = @()
$tempbuffer = @()
$questid = -1

### Process Paragraphs and store them in $Buffer, store content per question in array in $Buffer ###
  # Issue: this for loop can be combined with the next one if you find the time ;)
for ( $i=0; $i -lt $paragraphs.count; $i++ ) {
 
  # write-host "starting round $($i)" # Turn on for Debugging

  if ( !(Like $paragraphs[$i].text $Selector.question) ) {
    
    if ( ($paragraphs[$i].Pictures).count -like 1 ) {
      $tempbuffer += $imageFolder + $paragraphs[$i].Pictures.FileName
    }
    elseif ( Like $paragraphs[$i].Text $Selector.filter ) {
      # skip it
    }  
    else {
      $tempbuffer += $paragraphs[$i].text
    }
  }
  elseif ( (Like $paragraphs[$i].text $Selector.question) ) {
    $questionIndex += $i
    $questid ++
    $buffer.add($questid,$tempbuffer)
    $tempbuffer = @()
  } 
} # End for loop
# Access like this: "$buffer[questionnumber][indexnumberofcontentinquestion]""


### Process Buffer, store all the Question parts per Question in Objects, store Objects in $QuestionArray ###
for ( $i=0; $i -lt $buffer.count; $i++ ) { 
  $textExplanation = $false
  $QuestionArray += NewQuestion # Add new empty Question Object to array

  for ( $ii=0; $ii -lt $buffer[$i].count; $ii++ ) { # process parts of questions
    if ( $buffer[$i][$ii].length -lt 3 ) {
      # skip it
    }
    elseif ( Like $buffer[$i][$ii] $Selector.options ) { # Add part to Options Array
      $buffer[$i][$ii] -match $reg | out-null
      $QuestionArray[$i].answers += ($Matches[1] + ":)" + $Matches[2])
    }
    elseif ( Like $buffer[$i][$ii] $Selector.correct ) { # Add part to Correct Array
      $QuestionArray[$i].correct += $buffer[$i][$ii].Replace(" Answer:", ":")
    }
    elseif ( Like $buffer[$i][$ii] $Selector.imageFormat ) { # Add to image property
      $QuestionArray[$i].image += $buffer[$i][$ii]
    }
    elseif ( $textExplanation ) { # Add to Explanation Array
      $QuestionArray[$i].explanation += $buffer[$i][$ii]
    }
    elseif ( Like $buffer[$i][$ii] $Selector.explanation ) { # Add to explanation property
      $QuestionArray[$i].explanation += $buffer[$i][$ii].Insert(0, "Sol: ")
      $textExplanation = $true # Ensures all in-question-buffer is stored in Explanation array.
    }
    else { # Add to text array
      $QuestionArray[$i].text += $buffer[$i][$ii]
    }
  } # End of process parts of questions
} # End of process questions
 

### Write Question Parts in this order to Word File ###
for ( $i=1; $i -lt $QuestionArray.count; $i++ ) {
  # write-host "Add-Word: Starting round $($i)." # Turn on for Debugging

  Add-WordText -WordDocument $WordDocument -Text "Q:$($i))" -Supress $true  # Question Number
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true # Empty line
  # Question Text 
  $QuestionArray[$i].text | ForEach-Object {
    Add-WordText -WordDocument $WordDocument -Text "$($_)" -Supress $true
  }
  # Question Image
  if ($QuestionArray[$i].image.Length -gt 1) {
    $QuestionArray[$i].image | ForEach-Object {
      Add-WordPicture -WordDocument $WordDocument -ImagePath ("$($_)") -Supress $true
    }
  }
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true # Empty line
  # Question Options
  $QuestionArray[$i].answers | ForEach-Object {
    Add-WordText -WordDocument $WordDocument -Text "$($_)" -Supress $true
  }
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true # Empty line
  # Question Correct
  $QuestionArray[$i].correct | ForEach-Object {
    Add-WordText -WordDocument $WordDocument -Text "$($_)" -Supress $true
  }
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true # Empty line
  # Question Explanation
  $QuestionArray[$i].explanation | ForEach-Object {
    Add-WordText -WordDocument $WordDocument -Text "$($_)" -Supress $true
  }
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true # Empty line
} # End of Write Question Parts for loop

# Save changes to New WordDocument
Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument
