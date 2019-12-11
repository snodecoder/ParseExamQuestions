$WordFile = "C:\Codeprojects\ParseWordDocument\test.docx"
$QuestionSelector = "QUESTION"
$ExplanationSelector = "Explanation"
$OptionASelector = "A."
$OptionBSelector = "B."
$OptionCSelector = "C."
$OptionDSelector = "D."
$OptionESelector = "E."
$OptionFSelector = "F."
$OptionGSelector = "G."
$OptionHSelector = "H."

# Datastructure
function QuestionConstructor ([string]$id,[string]$text ){

  $textArray = @()

  $propertylist = @{
    id = $id
    ;text = $text
    ;answers = @()
    ;correct = ""
    ;explanation = ""
  }
  $question = New-Object psobject -Property $propertylist
  $question
}



$word = New-Object -ComObject "Word.Application"
$word.Visible = $true
$document = $word.Application.Documents.open($WordFile)
$paragraphs = $document.Paragraphs


# buffer
$buffer = @{}
$tempbuffer = @()
$questionIndex = @()
$questid = 0

# Create question index
for ($i=1; $i -lt 500; $i++) {
  write-host "starting round $($i)"
  if (!$paragraphs[$i].range.text.Contains($QuestionSelector)) {
    $tempbuffer += $paragraphs[$i].range.text

  }
  elseif ($paragraphs[$i].range.text.Contains($QuestionSelector)) {
    $questionIndex += $i
    $questid ++
    $buffer.add($questid,$tempbuffer)
    $tempbuffer = @()
  } 
} # End for loop

# Access buffer like this
# $buffer[questionnumber][indexnumberofcontentinquestion] 