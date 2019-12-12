$OldWord = "test.docx"
$folderPath = "C:\Codeprojects\ParseWordDocument\"
$imageFolder = "C:\Codeprojects\ParseWordDocument\test\word\media\"

$Selector = New-Object psobject -Property @{
  question = "QUESTION*"
  ;explanation = "https*"
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
    "*gratisexam"
  )
}



# install PSWriteWord module to easier edit word document: "install-module -name PSWriteWord -Force"
Import-Module PSWriteWord -Force

# how to use PSWriteWord Module
# $WordDocument = New-WordDocument $FilePath
# Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true
# Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -Verbose
# Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument

# Functions
function NewQuestion (){ # Create new question object

  $propertylist = @{
    text = @()
    ;image = ""
    ;answers = @()
    ;correct = @()
    ;explanation = ""
  }
  $question = New-Object psobject -Property $propertylist
  $question
} # End of function NewQuestion

function Like ( $str, $patterns ) {
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

# buffer
$buffer = @{}
$QuestionArray = @()
$questionIndex = @()
$tempbuffer = @()
$questid = -1


# Create question index 
# Access like this: "$buffer[questionnumber][indexnumberofcontentinquestion]""
for ( $i=0; $i -lt $paragraphs.count; $i++ ) {
  
  write-host "starting round $($i)"

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



for ( $i=0; $i -lt $buffer.count; $i++ ) { # process questions

  # add new empty Question Object to array
  $QuestionArray += NewQuestion

  for ( $ii=0; $ii -lt $buffer[$i].count; $ii++ ) { # process parts of questions
    if ( $buffer[$i][$ii].length -lt 3 ) {
      # skip it
    }
    elseif ( Like $buffer[$i][$ii] $Selector.options ) { # Add part to Options Array
      $QuestionArray[$i].answers += $buffer[$i][$ii].Replace(".", ":)")
    }
    elseif ( Like $buffer[$i][$ii] $Selector.correct ) { # Add part to Correct Array
      $QuestionArray[$i].correct += $buffer[$i][$ii].Replace(" Answer:", ":")
    }
    elseif ( Like $buffer[$i][$ii] $Selector.imageFormat ) { # Add to image property
      $QuestionArray[$i].image += $buffer[$i][$ii]
    }
    elseif ( Like $buffer[$i][$ii] $Selector.explanation ) { # Add to explanation property
      $QuestionArray[$i].explanation += $buffer[$i][$ii].Insert(0, "Sol: ")
    }
    else { # Add to text array
      $QuestionArray[$i].text += $buffer[$i][$ii]
    }
  } # End of process parts of questions
  # Continueing with next question
} # End of process questions
 

for ( $i=0; $i -lt $QuestionArray.count; $i++ ) {

  write-host "Add-Word: Starting round $($i)."
  # Question Header
  Add-WordText -WordDocument $WordDocument -Text "Q:$($i))" -Supress $true

  # Empty line
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true

  # 1. Question Text 
  $QuestionArray[$i].text | ForEach-Object {
    
    Add-WordText -WordDocument $WordDocument -Text "$($_)" -Supress $true
  }
  
  if ($QuestionArray[$i].image.Length -gt 1) {
    # 2. Question Image
    Add-WordPicture -WordDocument $WordDocument -ImagePath ("$($QuestionArray[$i].image)") -Supress $true

  }

  # Empty line
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true

  # 3. Question Options
  $QuestionArray[$i].answers | ForEach-Object {
    Add-WordText -WordDocument $WordDocument -Text "$($_)" -Supress $true
  }

  # Empty line
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true

  # 4. Question Correct
  $QuestionArray[$i].correct | ForEach-Object {
    Add-WordText -WordDocument $WordDocument -Text "$($_)" -Supress $true
  }
  
  # Empty line
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true

  if ($QuestionArray[$i].explanation.Length -gt 1) {
    # 5. Question Explanation
    Add-WordText -WordDocument $WordDocument -Text "$($QuestionArray[$i].explanation)" -Supress $true
  }

  # Empty line
  Add-WordText -WordDocument $WordDocument -Text " " -Supress $true

}




# Save changes to New WordDocument
Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument