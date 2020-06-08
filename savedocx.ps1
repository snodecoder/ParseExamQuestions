
Param(
  [array]$text,
  [string]$name
)

$folderPath = "C:\Codeprojects\ParseWordDocument\"
$word = New-Object -ComObject word.application
$word.Visible = $false
$doc = $word.documents.add()
$margin = 36 # 1.26 cm
$doc.PageSetup.LeftMargin = $margin
$doc.PageSetup.RightMargin = $margin
$doc.PageSetup.TopMargin = $margin
$doc.PageSetup.BottomMargin = $margin
$selection = $word.Selection



	


	$selection.TypeText($folderPath)
	$selection.TypeParagraph()

	$selection.TypeText($text);
	$selection.TypeParagraph()



$outputPath = $folderPath + $name
$doc.SaveAs($outputPath)
$doc.Close()
$word.Quit()
