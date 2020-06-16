
##### DEFINITIONS #####
class TextVariant # Text Vvariant (Large, Normal, Url)
{
  # Properties
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
  # Properties
  [string] $label
  [string] $Text


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

  # Constructor
  Question()
  {
    $this.variant # question variant
    $this.question # body of actual question
    $this.choices # body of actual choices
    $this.answer # array with true/false for every choice
    $this.explanation # explanation
  }
}

$exam = $null

$exam = [PSCustomObject]@{
  id = [int]$null # exam description
  title = [string]$null # exam description
  description = [string]$null # exam description
  author = [PSCustomObject]@{
    id = [int]$null # author ID
    name = [string]$null # author name
    image = [string]$null # author image
  }
  code = [string]$null # exam number
  time = [int]$null # maximum exam time
  pass = [int]$null # minimum score required to pass exam
  image = [string]$null # cover image of exam
  cover = [array[]] @() # fill array with addText method
  test = [array[]] @() # stores questions via addQuestion method
}

function AddTextVariant () { # Helper function to add textVariant blocks
  param(
    [Parameter(Mandatory=$true,
    HelpMessage="0=Image URL, 1=Normal Size, 2=Large Size")]
    [ValidateSet("ImageURL" , "Normal", "Large")]
    [string]$variant,
    [Parameter(Mandatory=$true,
    HelpMessage="Enter Text")]
    [string]$text
  )
  [int]$intVariant
  switch ($variant) {
    ImageURL { $intvariant = 0 }
    Normal { $intvariant = 1 }
    Large { $intvariant = 2 }
    Default {}
  }
  [TextVariant]::new($intVariant, $text)
}

function AddTextLabel () { # Helper function to add textLabel blocks
  param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("A" , "B", "C", "D", "E", "F", "G", "H", "I", "J", "K")]
    [string]$label,
    [Parameter(Mandatory=$true,
    HelpMessage="Enter Text")]
    [string]$text
  )
  [TextLabel]::new($label, $text)
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



### testing

$exam.test += [Question]::new()
$exam.test[0].answer += $true, $false
$exam.test[0].question += AddTextVariant -variant Normal -text "sdfsd" 
$exam.test[0].choices += AddTextLabel -label A -text "Voer een ip addres in"
$exam.test[0].variant = addQuestionType -type MultipleChoice


$exam | ConvertTo-Json -Depth 4 -Compress | Test-Json











$temp = [PSCustomObject]@{
  variant = [int]$null # defines the type of question (add via addQuestionType method)
  question = [array[]] @() # stores question parts (add via addTextVariant method)
  choices = [array[]] @() # stores answer choices (add via addChoice method)
  answer = [array[]] @() # stores correct and incorrect answers (add via addAnswers method)
  explanation = [array[]] @() # stores answer explanation (add via addTextVariant)
}





  # Add addText method to Exam object
$exam | Add-Member -Name addTextVariant -MemberType ScriptMethod -Value {
  param(
    [Parameter(Mandatory=$true,
    HelpMessage="0=Image URL, 1=Normal Size, 2=Large Size")]
    [int]$textSize,
    [Parameter(Mandatory=$true,
    HelpMessage="Enter Text")]
    [string]$text,
    [Parameter(Mandatory=$true,
    HelpMessage="Location to add text, default=cover",
    ValueFromPipeline=$true)]
    [ValidateSet("cover", "test.question", "test.explanation")]
    [psobject]$location
  )

  $this.$location += [PSCustomObject]@{ # Add object with values from input
    variant = [int]$textSize
    text = [string]$text
  }

} -Force

$exam.test[0].GetType()
$exam.test += $temp
$exam.addText(1, "test", "test.question")
$exam.test.gettype()

$exam | ConvertTo-Json | Test-Json



[string]$string = "test.question"



$exam.$string +=[PSCustomObject]@{ # Add object with values from input
  variant = [int]$textSize ="1"
  text = [string]$text = "asd"
}



Get-Content "test/jsonConfigFile.json" -Raw | Test-Json

$json = [ordered]@{}

(Get-Content "test/jsonConfigFile.json" -Raw | ConvertFrom-Json).PSObject.Properties |
    ForEach-Object { $json[$_.Name] = $_.Value }


function NewJsonExam () {
  New-Object [PSCustomObject] -Property @{
    id = [int] # exam description
    ;title = [string] # exam description
    ;description = [string] # exam description
    ;$author = [ordered] @{
      id = [int] # author ID
      ;name = [string] # author name
      ;image = [string] # author image
    }
    ;code = [string] # exam code
    ;time = [int] # exam time
    ;pass = [int] # minimum exam score needed
    ;image = [string] # exam cover image
    ;cover = [array[]] @() # takes objects generated by variant function
    ;test = [array[]] @() # takes objects generated by question function
  }
} # End of function newJsonExam

function NewJsonQuestion () {
  New-Object [PSCustomObject] -Property @{
    variant = [int] # question variant
    ;question = [array[]] @() # body of actual question
    ;choices = [array[]] @() # body of actual choices
    ;answer = [array[]] @() # array with true/false for every choice
    ;explanation = [array[]] @() # explanation
  }
} # End of function newJsonQuestion