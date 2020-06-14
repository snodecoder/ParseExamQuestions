Get-Content "test/jsonConfigFile.json" -Raw | Test-Json


$json = [ordered]@{}

(Get-Content "test/jsonConfigFile.json" -Raw | ConvertFrom-Json).PSObject.Properties |
    ForEach-Object { $json[$_.Name] = $_.Value }


$newJson = $json | ConvertTo-Json

$newJson | Test-Json

Import-Module .\functions.psm1 -Force

(Get-Content -Path "test/jsonConfigFile.json" | Measure-Object
$test = @()
$test = [array[]] @()
$test.GetType()
$test = NewJsonExam