
Param(
	[String]$temp
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

	$selection.TypeText($temp);
	$selection.TypeParagraph()



$outputPath = $folderPath + "sources.docx"
$doc.SaveAs($outputPath)
$doc.Close()
$word.Quit()
