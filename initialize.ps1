class Question # Question constructor
{
  [string]$index
  [string]$topic
  [string]$type
  [string]$question
  [string]$option0
  [string]$option1
  [string]$option2
  [string]$option3
  [string]$option4
  [string]$option5
  [string]$option6
  [string]$option7
  [string]$explanation
  [string]$answer0
  [string]$answer1
  [string]$answer2
  [string]$answer3
  [string]$score0
  [string]$score1
  [string]$score2
  [string]$score3

  Question() # Constructor
  {
    $this.index
    $this.topic
    $this.type
    $this.question
    $this.option0
    $this.option1
    $this.option2
    $this.option3
    $this.option4
    $this.option5
    $this.option6
    $this.option7
    $this.explanation
    $this.answer0
    $this.answer1
    $this.answer2
    $this.answer3
    $this.score0
    $this.score1
    $this.score2
    $this.score3
  }
}

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


# Add the working directory to the environment path.
# This is required for the ChromeDriver to work.
$workingPath = "$PSScriptRoot/selenium"
Write-Host "Set envPath to $workingPath"
if (($env:Path -split ';') -notcontains $workingPath) {
    $env:Path += ";$workingPath"
}

Import-Module $workingPath\WebDriver.dll

$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver
$chromeDriver.manage().timeouts().implicitWait = [System.TimeSpan]::FromSeconds([int]3)

function Get-SeleniumElements($xPath, $className) {
    try {
        if ($xPath.length -gt 0) {          $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPath($xPath)) }
        elseif ($className.length -gt 0) {  $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::ClassName($className)) }

        if ($result.count -gt 0) { return $result}
        else { return $null}
    }
    catch{
        Write-Error $_
    }
}

function Click-SeleniumElementButton ($xPath, $className) {
    try {
        if ($xPath.length -gt 0) {          $result = $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath($xPath)).click() }
        elseif ($className.length -gt 0) {  $result = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ClassName($className)).click() }

        if ($result.count -gt 0) { return $result}
        else { return $null}
    }
    catch{
        Write-Error $_
    }
}

function Get-SeleniumElementsText ($xPath, $className) {
    try {
        if ($xPath.length -gt 0) {          $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPath($xPath)).Text }
        elseif ($className.length -gt 0) {  $result = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ClassName($className)).Text }

        if ($result.count -gt 0) { return $result}
        else { return $null}
    }
    catch{
        Write-Error $_
    }
}

function Get-SeleniumElementOuterHTML ($xPath, $className) {
    try {
        if ($xPath.length -gt 0) {          $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPath($xPath)).GetAttribute('outerHTML') }
        elseif ($className.length -gt 0) {  $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::ClassName($className)).GetAttribute('outerHTML') }

        if ($result.count -gt 0) { return $result}
        else { return $null}
    }
    catch{
        Write-Error $_
    }
}

function Get-SeleniumElementInnerHTML ($xPath, $className) {
    try {
        if ($xPath.length -gt 0) {          $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPath($xPath)).GetAttribute('innerHTML') }
        elseif ($className.length -gt 0) {  $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::ClassName($className)).GetAttribute('innerHTML') }

        if ($result.count -gt 0) { return $result}
        else { return $null}
    }
    catch{
        Write-Error $_
    }
}
function Get-SeleniumElementAttribute ($xPath, $className, $attribute) {
    try {
        if ($xPath.length -gt 0) {          $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPath($xPath)).GetAttribute($attribute) }
        elseif ($className.length -gt 0) {  $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::ClassName($className)).GetAttribute($attribute) }

        if ($result.count -gt 0) { return $result}
        else { return $null}
    }
    catch{
        Write-Error $_
    }
}


function Get-SeleniumElementHref ($xPath, $className) {
    try {
        if ($xPath.length -gt 0) {          $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPath($xPath)).GetAttribute('href') }
        elseif ($className.length -gt 0) {  $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::ClassName($className)).GetAttribute('href') }

        if ($result.count -gt 0) { return $result}
        else { return $null}
    }
    catch{
        Write-Error $_
    }
}

function Get-SeleniumElementSrc ($xPath) {
    try {
        if ($xPath.length -gt 0) {          $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPath($xPath)).GetAttribute('src') }
        elseif ($className.length -gt 0) {  $result = $ChromeDriver.FindElements([OpenQA.Selenium.By]::ClassName($className)).GetAttribute('src') }

        if ($result.count -gt 0) { return $result}
        else { return $null}
    }
    catch{
        Write-Error $_
    }
}

function Get-SeleniumElementChildren ($xPath, $ClassName) {
    try {
        if ($xPath.length -gt 0) {          $element = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPath($xPath)) }
        elseif ($className.length -gt 0) {  $element = $ChromeDriver.FindElements([OpenQA.Selenium.By]::ClassName($className)) }

        if ($element.Count -gt 0) {         $result = $element.FindElements([OpenQA.Selenium.By]::XPath(".//*")) }

        if ($result.count -gt 0) { return $result }
        else { return $null}
    }
    catch{
        Write-Error $_
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
