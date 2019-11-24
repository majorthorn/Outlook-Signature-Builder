$path = (Get-Location).path
$objWord = New-Object -ComObject Word.Application
$objWord.Visible = $true
$objDoc = $objWord.Documents.Open("$path\Test.docx")
$objSelection = $objWord.Selection

#region
$ReplaceAll = 2
$FindContinue = 1
$MatchCase = $False
$MatchWholeWord = $True
$MatchWildcards = $False
$MatchSoundsLike = $False
$MatchAllWordForms = $False
$Forward = $True
$Wrap = $FindContinue
$Format = $False
#endregion

$FindText = "[ Test ]"
$ReplaceText = "Mary had a little lamb"

$objSelection.Find.Execute($FindText,$MatchCase,
  $MatchWholeWord,$MatchWildcards,$MatchSoundsLike,
  $MatchAllWordForms,$Forward,$Wrap,$Format,
  $ReplaceText,$ReplaceAll)

$objDoc.Save()
$objDoc.Close()
$objWord.Quit()