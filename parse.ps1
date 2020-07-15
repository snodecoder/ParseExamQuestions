$OldWord = "742.docx"
$folderPath = "C:\Codeprojects\ParseWordDocument\"
$mediaFolder = "C:\Codeprojects\ParseWordDocument\docx\word\media\"
$imageFolder = $folderPath + "images\"
$Outputformat = "CSV" # Enter 'CSV' or 'Word' if you would like CSV or DOCX output format
$examNumber = "70-742"
$CSVFormat = "QuestionType;Question;Description;correct option number;Option 1;Option 2;Option 3;Option 4;Option 5;Option 6;Option 7;Option 8;Option 9;Option 10;Option 11;Option 12;"

$reg = '([A-Z]{1})[\.](.*)' # Regex match string to select First letter in Option, replace '.' with ':)', finally add answer.
$Selector = New-Object psobject -Property @{
  question = "NEW QUESTION*"
  ;explanation = "Explanation*"
  ;correct = "Answer*"
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
    type = @()
    ;text = @()
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

function ConvertAnswer($_) {
  $input = $_.tostring()
  switch ( $input ) {
    "A" {"1"; Break}
    "B" {"2"; break}
    "C" {"3"; break}
    "D" {"4"; break}
    "E" {"5"; break}
    "F" {"6"; break}
    "G" {"7"; break}
    "H" {"8"; break}
    "I" {"9"; break}
    "J" {"10"; break}
    "K" {"11"; break}
    "L" {"12"; break}
  }
} # End of function ConvertAnswer

# Prepare Word Document
$WordDocument = New-WordDocument ($folderPath + "new.docx")
$OldWordDocument = Get-WordDocument -FilePath ($folderPath + $OldWord)
$paragraphs = $OldWordDocument.Paragraphs

# Create Image folder (for exported images) in working directory, if it not already exists
 if ( (Test-Path -Path ($imageFolder)) -like "False" ) {
   New-Item -Path $folderPath -Name "images" -ItemType Directory
 }

# Prepare data structure
$buffer = @{}
$QuestionArray = @()
$tempbuffer = @()
$questid = -1


### Process Paragraphs and store them in $Buffer, store content per question in array in $Buffer ###
# Access like this: "$buffer[questionnumber][indexnumberofcontentinquestion]""
for ( $i=0; $i -lt $paragraphs.count; $i++ ) {
  # write-host "starting round $($i)" # Turn on for Debugging
  if ( !(Like $paragraphs[$i].text $Selector.question) ) {
    
    if ( ($paragraphs[$i].Pictures).count -like 1 ) {
      #$tempbuffer += ($examNumber + "_" + $paragraphs[$i].Pictures.FileName)
      $tempbuffer += $paragraphs[$i].Pictures.FileName
      write-host $i
      Copy-Item -Path ($mediaFolder + $paragraphs[$i].Pictures.FileName) -Destination ($imageFolder + $examNumber + "_" + $paragraphs[$i].Pictures.FileName) -ErrorAction Ignore # Copy image to export folder for upload to server
    }
    elseif ( Like $paragraphs[$i].Text $Selector.filter ) {
      # skip it
    }  
    else {
      $tempbuffer += $paragraphs[$i].text
    }
  }
  elseif ( (Like $paragraphs[$i].text $Selector.question) ) {
    $questid ++
    $buffer.add($questid,$tempbuffer)
    $tempbuffer = @()
  } 
} # End for loop


############### Store all the Question parts per Question in Objects, store Objects in $QuestionArray ###############

### Process Buffer to WORD
if ($Outputformat -like "Word") {
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
  Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
  
} # End of Process buffer to Word

### Process buffer to CSV
elseif ($Outputformat -like "CSV") {
  for ( $i=0; $i -lt $buffer.count; $i++ ) { 

    $textExplanation = $false
    $QuestionArray += NewQuestion # Add new empty Question Object to array

    for ( $ii=0; $ii -lt $buffer[$i].count; $ii++ ) { # process parts of questions
      if ( $buffer[$i][$ii].length -lt 3 ) {
        # skip it
      }
      elseif ( Like $buffer[$i][$ii] $Selector.options ) { # Add part to Options Array
        $buffer[$i][$ii] -match $reg | out-null
        $QuestionArray[$i].answers += $Matches[2].Trim(" ")
      }
      elseif ( Like $buffer[$i][$ii] $Selector.correct ) { # Add part to Correct Array
        $tempCorrect = $buffer[$i][$ii].Replace( ($Selector.correct.Trim("*") + ": "), "" )

        if ($tempCorrect.Length -like 1) { # convert single answer to decimal
          $QuestionArray[$i].correct += ConvertAnswer $tempCorrect
          $QuestionArray[$i].type = "0" # Set MultipleChoice question type to single answer
        }
        elseif ($tempCorrect.Length -gt 1) { # Convert multiple answers to decimal
          $tempSplit = $tempCorrect.ToCharArray() | ForEach-Object {$_.tostring()}
          $QuestionArray[$i].correct += [system.String]::Join(",", ($tempSplit | ForEach-Object { ConvertAnswer $_ } ))
          $QuestionArray[$i].type = "1" # Set MultipleChoice question type to multiple answer
        }
      }
      elseif ( Like $buffer[$i][$ii] $Selector.imageFormat ) { # Add to image property
        write-host $i
        $QuestionArray[$i].text += ("<img src=../../../../upload/images/$($buffer[$i][$ii])" + " />")
      }
      elseif ( $textExplanation ) { # Add to Explanation Array
        $QuestionArray[$i].explanation += $buffer[$i][$ii] + "<br><br>"
      }
      elseif ( Like $buffer[$i][$ii] $Selector.explanation ) { # Add to explanation property
        $QuestionArray[$i].explanation += $buffer[$i][$ii]  + "<br><br>"
        $textExplanation = $true # Ensures all in-question-buffer is stored in Explanation array.
      }
      else { # Add to text array
        $QuestionArray[$i].text += $buffer[$i][$ii] + "<br><br>"
      }
    } # End of process parts of questions
  } # End of process questions

  # Save Data as CSV
  $prepareCsv = for ($i=1; $i -lt $QuestionArray.count; $i++) {
    
    if ($QuestionArray[$i].correct.count -like 0) {
      # skip it
    }
    else {
      $QuestionArray[$i].type + ";" + $QuestionArray[$i].text + [system.String]::Join("<br>", $QuestionArray[$i].image) + ";" + $QuestionArray[$i].explanation + ";" + $QuestionArray[$i].correct + ";" + [system.String]::Join(";", ($QuestionArray[$i].answers) ) + ";"
    }

  } # End prepare CSV

  $CSVFormat | Out-File -FilePath ($folderPath + "test.csv") -Encoding utf8 -Force
  $prepareCsv | Out-File -FilePath ($folderPath + "test.csv") -Encoding utf8 -Append

} # End of Process Buffer to CSV

