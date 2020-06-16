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


### Functions ###

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

function NewJsonExam () {
  [PSCustomObject]@{
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
} # End of function newJsonExam


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

  #Get-ChildItem -Path ($wordFile.BaseName + "\word\media\") | ForEach-Object {
  #  Copy-Item -Path ($wordFile.BaseName + "\word\media\*") -Destination ($folderPath + "\images")
  #}
  $zipFile = Get-ChildItem -Path ($folderPath + $wordFile.BaseName + ".zip") -Filter *.zip 
  Rename-Item -Path $zipFile.FullName -NewName ($zipFile.BaseName + ".docx") 
  #Remove-Item -Path ($folderPath + "\" + $zipFile.BaseName) -Recurse
} # End of function extractWordImages

function NewQuestion (){ # Create new question object
  $propertylist = [ordered] @{
    index = [string]
    ;section = @()
    ;type = @()
    ;text = @()
    ;image = @()
    ;answers = @()
    ;correct = @()
    ;explanation = @()
  }
  $question = New-Object psobject -Property $propertylist
  $question
} # End of function NewQuestion


Export-ModuleMember -Function NewQuestion, Like, ConvertAnswer, booleanAnswer, insertChoice, NewJsonExam, AddTextVariant, AddTextLabel, addQuestionType, ExtractWordImages
