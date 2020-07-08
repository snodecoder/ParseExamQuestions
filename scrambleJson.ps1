$folderPath = "C:\Codeprojects\ParseWordDocument\"
$examCode = "70-742"
$maxQuestions = 50

$exam = Get-Content -Raw -Path ($folderPath + $examCode + ".json")| ConvertFrom-Json -Depth 5
[System.Collections.ArrayList]$shuffled = $exam.test | Sort-Object {Get-Random}

$exam.test = @() # delete questions from Exam
$jsonExam = $null

function copyQuestions() { # Copy and remove questions

  # If set maxQuestions to shuffled.count if lower than maxQuestions
  if ($shuffled.Count -lt $maxQuestions) {
    $maxQuestions = $shuffled.Count
  }

  for ($i=0; $i -lt $maxQuestions; $i++) {
    $exam.test += $shuffled[$i]
  }

  # Remove copied questions
  if ($shuffled.Count -lt 50) {
    $shuffled.RemoveRange(0, $shuffled.count)
  }
  else {
    $shuffled.RemoveRange(0, 50)
  }

}

$index = 1 # numberIndex to increment filename

while ($shuffled.Count -gt 0) {

  write-host "exam.test.coumt: "
  $exam.test.Count
  Write-Host "shuffled.count: "
  $shuffled.count

  # Copy Questions to $jsonExam, convert to JSON
  copyQuestions
  $exam.description = "Variant: $($index) | $($exam.test.Count) questions available in Multiple Choice en Multiple Answer format."
  $exam.cover[1] = "Variant: $($index) | $($exam.test.Count) questions available in Multiple Choice en Multiple Answer format."
  $exam.image = "https://pbs.twimg.com/profile_images/1080479698742902784/724c4osq_400x400.jpg"

  $jsonExam = $exam | ConvertTo-Json -Depth 5

  if ( $jsonExam | Test-Json ) {
    $jsonExam | Out-File -FilePath ($folderPath + "$($examCode)_$($index).json") -Force
    Write-Host "Exported $($exam.test.Count) questions to JSON file :)" -ForegroundColor Green
  }
  else {
    Write-Warning "Please check generated jsonExam. It is not a valid JSON file."
  }

  # Increment filename index, clear array in $exam.test
  $index++
  $exam.test = @()
}

