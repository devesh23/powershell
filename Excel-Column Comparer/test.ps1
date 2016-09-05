$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true

$wb = $xl.Workbooks.Open("C:\Users\devesh23\Desktop\Compare.xlsx")
$ws = $wb.Sheets.Item(1)

$rows = $ws.UsedRange.Rows.Count
$cols = $ws.UsedRange.Columns.Count

$len=@{}

for($i=0 ; $i -lt $cols ; $i++) {
    $k=[char]($i+65)
  $len[$i]=($xl.WorksheetFunction.CountIf($ws.Range($k+"1:"+$k+$rows), "<>") - 1)
}
$len[0]
$wb.Close()
$xl.Quit()