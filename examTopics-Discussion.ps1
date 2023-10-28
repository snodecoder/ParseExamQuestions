
param (
  $examCode= "AZ-140"
  ,$examTitle = $examCode
  ,$examDescription = "Practice questions in Multiple Choice en Multiple Answer format."
  ,$examDuration = 120 # Maximum time for exam
  ,$examKeywords = "AVD, Azure, Fundamentals"
  ,[int] $TotalPages = 1379 # update this to the total number of pages found when opening $url_discussion
)

$DebugPreference = 'Continue'


. .\initialize.ps1

[array]$DiscussionLinks = @()
[array]$Exam = @()
[array]$ProcessQuestionsManually = @()
$logfile = ".\errorlog.txt"

# Launch a browser and go to URL
$url_base = "https://www.examtopics.com"
$url_discussion = "$url_base/discussions/microsoft/"

$ChromeDriver.Navigate().GoToURL($url_discussion)

# Login
$BTN_login = "/html/body/div[1]/div/div/div/div[2]/div/div[1]/ul/li[1]/a"
$ChromeDriver.FindElements([OpenQA.Selenium.By]::xPath($BTN_login)).click()
$FORM_user = '//*[@id="login-modal"]/div/div/div[2]/div/form/div[1]/div/input'
$FORM_pass = '//*[@id="login-modal"]/div/div/div[2]/div/form/div[2]/div[1]/input'
$FORM_login = '//*[@id="login-modal"]/div/div/div[2]/div/form/button'


[string]$username = "snodecoder"
[string]$password = "WM?ejHbpsn.49QV"
$ChromeDriver.FindElement([OpenQA.Selenium.By]::xPath($FORM_user)).SendKeys($username)
$ChromeDriver.FindElement([OpenQA.Selenium.By]::xPath($FORM_pass)).SendKeys($password)
$ChromeDriver.FindElement([OpenQA.Selenium.By]::xPath($FORM_login)).Click()

# ETA Calculation
$start = Get-Date

# Search Discussions for examcode, retrieve urls to questions
for ($page = 1; $page -le $TotalPages; $page++)
{
    # Progress Tracking
    $prct = $page / $TotalPages

    $elapsed = (Get-Date) - $start
    $totalTime = ($elapsed.TotalSeconds) / $prct
    $remain = $totalTime - $elapsed.TotalSeconds
    $eta = (Get-Date).AddSeconds($remain)

    # Display
    $activity = "Gathering relevant Discussion Links ETA"
    $status = ("$($prct.ToString('P')) % ($page/$TotalPages) {0:dd\.hh\:mm\:ss} eta $eta" -f (New-TimeSpan -seconds $remain))
    Write-Progress -Activity $activity -Status $status -PercentComplete ($prct * 100)

    # Operation
    $CLASS_discussion = "discussion-link"
    $Results = $ChromeDriver.FindElements([OpenQA.Selenium.By]::ClassName($CLASS_Discussion)) | Where-Object { $_.Text -like "*$($ExamCode)*" }

    # Store link to found discussion in array
    if ($Results.Count -gt 0)
    {
        # $Result = $Results[0] # for debugging
        foreach ($Result in $Results)
        {
            $link = $Result.GetAttribute('href')
            $DiscussionLinks += $link
            Write-Debug "Found: $link"
        }
    }

    # Continue to next page if available
    $BTN_NextPage = "/html/body/div[2]/div/div[6]/div/span/span[2]/a"
    if ($page -gt 1) { $BTN_NextPage = "/html/body/div[2]/div/div[6]/div/span/span[2]/a[2]" }

    $NextPage = $ChromeDriver.FindElements([OpenQA.Selenium.By]::XPATH($BTN_NextPage))



    if ($NextPage.count -eq 1) { $NextPage.click()}
    elseif ($nextPage.count -eq 0) { Read-Host -Prompt "This should be the end and all questions should be stored in Buffer. Press any key to continue"}

} # end for loop $page
Write-Host "Found: $($DiscussionLinks.Count) links to discussions for examcode: $($ExamCode)."





# ETA Calculation
$start = Get-Date
get-content -Path ".\az140-links.txt" | ForEach-Object { $discussionLinks += $_ }

 #$link = $DiscussionLinks[0] # for debug
for ($Page = 0; $page -lt $TotalPages; $page++)
{

    try {
        # Progress Tracking
        $prct = $page / $DiscussionLinks.Count

        $elapsed = (Get-Date) - $start
        $totalTime = ($elapsed.TotalSeconds) / $prct
        $remain = $totalTime - $elapsed.TotalSeconds
        $eta = (Get-Date).AddSeconds($remain)

        # Display
        $activity = "Crawling links to extract questions"
        $status = ("$($prct.ToString('P')) % ($page/$($DiscussionLinks.Count)) {0:dd\.hh\:mm\:ss} eta $eta" -f (New-TimeSpan -seconds $remain))
        Write-Progress -Activity $activity -Status $status -PercentComplete ($prct * 100)


        $QuestionObject = [Question]::new()
        $ChromeDriver.Navigate().GoToURL($DiscussionLinks[$page])

        # Proces question index
        $QuestionInfo = (Get-SeleniumElementText -xPath "/html/body/div[2]/div/div[3]/div/div[1]/div[1]/div").ReplaceLineEndings("`n").split("`n")
        $QuestionObject.index = $QuestionInfo[0] -replace '([^0-9])+'
        $QuestionObject.topic = $QuestionInfo[1] -replace '([^0-9])+'

        # Process Question Text
        $QuestionObject.question = Get-SeleniumElementAttribute -xPath "/html/body/div[2]/div/div[3]/div/div[1]/div[2]/p" -attribute "innerHTML"
        $QuestionObject.question = $QuestionObject.question -replace '([^a-zA-Z0-9!-~ ])'
        $QuestionObject.question = $QuestionObject.question.Replace("<img src=""","<img src=""$($url_Base)")

        # Process Options
        $QuestionOptions = Get-SeleniumElementChildren -ClassName "question-choices-container" | Where-Object { $_.TagName -eq 'li'}

        # Multiple Choice or Single Choice question
        if ($QuestionOptions.Count -gt 0)
        {
            for ($i=0; $i -lt $QuestionOptions.Length; $i++)
            {
                $Questionobject."option$i" = [regex]::Replace($QuestionOptions[$i].Text,'[a-zA-Z]\.', "").Replace("Most Voted", "").Replace("<img src=""","<img src=""$($url_Base)").Trim()
            }
        }


        # Process Correct Answer
        Click-SeleniumElementButton -className "reveal-solution"
        $CorrectAnswers = Get-SeleniumElementsText -className "correct-answer"

        if ($CorrectAnswers.Length -eq 0 -and $QuestionOptions.count -eq 0)
        {
            Write-warning "# Probably a Drag and Drop question, this type will need manual action. This question will be added to ProcessQuestionsManually array. $($QuestionObject.index), $link."
            $QuestionObject.type = "drap and drop"
        }
        elseif ($CorrectAnswers.Length -eq 1)
        {
            $QuestionObject.answer0 = $CorrectAnswers
            $QuestionObject.score0 = "1.0"
        }
        elseif ($CorrectAnswers.Length -gt 1 -and $CorrectAnswers.Length -le 4)
        {
            $QuestionObject.type = "multiple choice"

            for ($i=0; $i -lt 4; $i++)
            {
                $QuestionObject.answer0 = $CorrectAnswers
                $QuestionObject."score$($i)" = "1.0"
            }
        }
        else { Write-Warning "Possible error detected in Correct Answer for question: $($QuestionObject.index). The current correct answer is: $($CorrectAnswers), $link"}


        # Process Explanation / reference
        $QuestionObject.explanation = (Get-SeleniumElementText -className "answer-description") -split "(?=https?:)"

    }
    catch {
        Write-warning $_
        $Errormessage = "$(Get-Date) | QuestionIndex: $($QuestionObject.index) | Errormessage: $($Error[0]) | URL: $($link)"
        Write-Warning $Errormessage
        $Errormessage | out-file -FilePath $logfile -Append -Force
        $QuestionObject | Out-String | Out-File -FilePath $logfile -Append -Force
    }

    if ($QuestionObject.type -eq "drap and drop") { $ProcessQuestionsManually += $QuestionObject}
    else { $Exam += $QuestionObject}

}
    read-host -Prompt "press key to exit"
   # $buffer += "</body></html>"
   $exam.count


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
     return

   }

   ########## Convert Exam to CSV and Export it ##########
   $ColumnCount = ($exam | Get-Member -MemberType Property).count

   $header = "Title;$($examTitle)$(printSemiColon $ColumnCount `n)`
   Description;$($examDescription)$(printSemiColon $ColumnCount `n)`
   Duration;$($examDuration)$(printSemiColon $ColumnCount `n)`
   Keywords;$($examKeywords)$(printSemiColon $ColumnCount `n)"


   for($i=0; $i -lt $exam.Count; $i++) {
    $exam[$i].question = "<p>$($exam[$i].question)</p>"
   }



   $header | Out-File -FilePath ($folderPath + "$examCode.csv") -Force
   $CSV = $exam | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
   $CSV | Out-File -FilePath ".\$examCode.csv" -Append

# Close browser
 #$ChromeDriver.Quit()

