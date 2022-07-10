param (
  $htmlSourceFilePath = ".\exam\az104.html"
  ,$htmlImagesFilePath = ".\exam\az104_files\"
  ,$examCode= "AZ104"
  ,$examTitle = "AZ104"
  ,$examDescription = "Practice questions in Multiple Choice en Multiple Answer format."
  ,$examDuration = 120 # Maximum time for exam
  ,$examKeywords = "Azure, Fundamentals"
  ,$imageURLPrefix = "https://files.doorhetgeluid.nl/exams/$($examCode)/"
  ,$WordFileName = "$($examCode).docx"
  ,$folderPath = "C:\CodeProjects\ParseWordDocument\"
  ,$ImagePath = "Z:\IIS_Files\exams\"
)

if (! (Test-Path "$ImagePath$examCode") ) { New-Item -Path $ImagePath -Name $examCode -ItemType Directory }
if (Test-Path "$imagePath$examCode") { $imagePath = "$imagePath$examCode\"}

$Selector = New-Object psobject -Property @{
    question = "Question #*"
    
    ;correct = "Correct Answer:*"
    ;explanation = "Explanation:*"
    #;section = "- (Exam Topic*"
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
      "Note: This question is part of a series of questions that present the same scenario. Each question in the series contains a unique solution that might meet the stated goals. Some question sets might have more than one correct solution, while others might not have a correct solution.<BR>After you answer a question in this section, you will NOT be able to return to it. As a result, these questions will not appear in the review screen.<BR>"
      ,"âœ‘"
      ,"×’â‚¬`"&gt"
    )
    ;FilterType = @(
      "*hotspot*"
      ,"*drag drop*"
    )
    class = @(
        [pscustomobject]@{
            QuestionNumber = "exam-question-card"
            Topic = "question-title-topic"
            Question = "card-text"
            Options = "question-choices-container"
            Answer = "correct-answer"
            Explanation = "answer-description"
        }
    )
  } # End of Selector object


# HTML Class Config



$ErrorActionPreference = 'Stop'

# Function to convert the html COM object graph into PSCustomObjects
# This makes the tree a bit easier to work with since you can access by node name
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

function Remove-SpecialCharacters ( $string) {


  $charactersToKeep =@("/", "\","!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "+", "=", "``", "<", ">", ",", ".", "?", "{", "}", "[", "]", "|", ";", ":")
  [string]$ExcludeCharacters = foreach ($char in $charactersToKeep) {"/$char"}
  $ExcludeCharacters = $ExcludeCharacters.Replace(" ", "")
  $regex = '[^\p{L}\p{Nd}'
  $regexString = $regex + $ExcludeCharacters
  $string -replace $regexString, ''

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

function Export-Html  ($html)
{
try{
    if ($null -ne $html.className) { $className = $html.className }
    else { $className = "none"}

    if ($null -ne $html.innerHtml) { $innerHtml = $html.innerHtml }
    else { $innerHtml = "none"}

    if ($null -ne $html.textContent) { $textContent = $html.textContent }
    else { $textContent = "none"}

    if ($null -ne $html.nodeName) { $nodeName = $html.nodeName }
    else { $nodeName = "none"}

    
    if ($null -ne $html.className) { $Name = $html.className }
    elseif ( $null -ne $html.nodeName) {$Name = $Html.nodeName }
    else { $Name = "none"}

    [pscustomobject] @{
        innerHtml   = $innerHtml
        textContent = $textContent
        className   = $className
        nodeName    = $html.NodeName
    }
        
    
}
catch{
    Write-host "there was an error with input: $($html | select *)"
    Write-host $_
}
    
}

# Load HTML structure
$content = Remove-SpecialCharacters (Get-Content $htmlSourceFilePath -Encoding UTF8)


$page = New-Object -ComObject "HTMLFile"
$page.IHTMLDocument2_write($content)

$questions = $page.body.getElementsByclassname("exam-question-card") 



# Prepare data structure
[array]$exam = @()
$QuestionObject = [Question]::new()
$Question_index = 0
$tempOptions = $null



# $Question = $Questions[0]
try{
    
    foreach ($Question in $Questions) 
    {
        Write-Progress -Activity "questions" -PercentComplete ( $Question_index / $Questions.length * 100 )
        Write-Host "Starting Question: $Question_index" -ForegroundColor Green
        $QuestionObject = [Question]::new()

        [array]$QuestionParts = $question.all | foreach {Export-html $_}
  
        
        ##### ACTUAL QUESTION #####
        $text = $null
        $text = $question.getElementsByTagName("p") | Where-Object {$_.classname -eq $Selector.class.question } | select innerHtml -ExpandProperty innerHtml

        ##### Skip entire question if question is one of FilterType #####
        if (Like $text $selector.FilterType) {
          Write-Host 'Skiping' -ForegroundColor Green
          continue
        }

        ##### Images #####
        [array]$images = $question.getElementsByTagName("img") | select src 
                
        foreach ($image in $images) { # Copy images used to webserver share
            $image.src = $image.src.replace("about:", "")
            $imageFileName = $image.src.split("/") | select -Last 1
            Copy-Item -Path "$htmlImagesFilePath$imageFileName" -Destination "$ImagePath$imageFileName" -Force

            # Update Image url to a publicly available url.
            if ($text.contains($image.src)) { $text = $text.replace("$($image.src)`"", "$($imageURLPrefix + $imageFileName)`" style='max-width: 100%;' ")}
        }
        
        if ( !($text.StartsWith("<p>"))) { $text = $text.Insert(0, "<p>")}
        if ( !($text.EndsWith("</p>"))) { $text = $text + "</p>"}
        
        # Store result
        foreach ($filter in $Selector.filter) {
          if ($text.contains($filter)) { $text = $text.replace($filter, "")}
        }
        $questionobject.question += $text



        ##### Possible Answers ######
        $NumberofOptions = $question.getElementsByTagName("li").length
        
        for ($i=0; $i -lt $NumberofOptions; $i++) {
          $tempOptions += ($question.getElementsByTagName("li")[$i].childnodes[1].textcontent)
        }


        Foreach ($part in $QuestionParts) {
            ##### QUESTION Index #####
            if     ($part.textContent -like "*$($Selector.class.questionnumber)*") { $QuestionObject.no = $part.textContent}
            elseif ($part.innerHtml -like "*$($Selector.class.questionnumber)*") { $QuestionObject.no = $part.textContent}
            
            ##### ACTUAL QUESTION #####
            elseif ($part.className -eq $Selector.class.question) { $questionobject.question += $part.textContent}
          
            ##### CORRECT ANSWERs #####
            elseif ($part.className -eq $Selector.class.answer) { $QuestionObject.answer1 = $part.textContent }

            ##### EXPLANATION #####
            elseif ($part.className -like $selector.class.explanation) { $QuestionObject.explanation = $part.innerHtml}
        }

        
        for ($ii = 0; $ii -lt $tempOptions.Count; $ii++) {
            $item = "option$($ii+1)"
            $QuestionObject.$item = $tempOptions[$ii]
        }


        # add extra points when multiple answers are corect
        $numberAnswers = $QuestionObject.answer1.Length
        for ($ii = 1; $ii -le $numberAnswers; $ii++) {
            $item = "score$($ii)"
            $QuestionObject.$item = "1"            
        }
        if ($numberAnswers -gt 1) { $QuestionObject.type = "multiple choice"}


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
        
        # filter unwanted textparts
        foreach ($filter in $Selector.filter) {
          if ($QuestionObject.question.contains($filter)) { $QuestionObject.question = $QuestionObject.question.replace($filter, "")}
          if ($QuestionObject.explanation.contains($filter)) { $QuestionObject.explanation = $QuestionObject.explanation.replace($filter, "")}
        }

        
        # store questionobject in Exam object
        if ($QuestionObject.question.length -gt 0) {
            $exam += $QuestionObject
        }

        # Reset for next question
        $tempOptions = @()
        $Question_index++
    }
    Write-Host "Finished processing document." -ForegroundColor Green
}

catch{
  write-warning $_
  write-host $Error[0].ScriptStackTrace 
  Write-Host "Question summary: " -ForegroundColor Blue
  $QuestionObject
  Write-Host "Question: Text" -ForegroundColor Blue
  $QuestionObject.question | Format-Table
  Write-Host "Question: Choices" -ForegroundColor Blue
  $QuestionObject.choices | Format-Table
  Write-Host "Question: Answers" -ForegroundColor Blue
  $QuestionObject.answer | Format-List
  Write-Host "Question: Explanation" -ForegroundColor Blue
  $QuestionObject.explanation

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






$header | Out-File -FilePath ($folderPath + "exam.csv") -Force
$CSV = $exam | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
$CSV | Out-File -FilePath ($folderPath + "exam.csv") -Append
$cs

#$CSV | Export-Csv -path ($folderPath + "$($examCode).CSV") -Delimiter ";" -UseQuotes Never -Encoding unicode


#  $CSV | Out-File -FilePath ($folderPath + "$($examCode).CSV") -Force
Write-Host "Exported $($exam.Count) questions to CSV file :)" -ForegroundColor Green

Write-Host "Open csv in Excel (without loading the csv via a query) and review if everything is in order, then save in.xls format, and import in MTestM." -ForegroundColor Green

