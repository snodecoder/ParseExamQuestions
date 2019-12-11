$WordFile = "C:\Codeprojects\ParseWordDocument\test1page.docx"
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
$buffer = @()
$questionIndex = @()

# Create question index
for ($i=1; $i -lt $paragraphs.count; $i++) {
  if ($paragraphs[$i].range.text.Contains($QuestionSelector)) {
    $questionIndex += $i
  } 
} # End for loop

for ($i=1; $i -lt $paragraphs.count; $i++) {
  if ($i -notin $questionIndex ) {
    $buffer += $paragraphs[$i].range.text
  }
  #$buffer += $_.range.text


} # End for loop

$buffer