."$home\desktop\Excel-Column Comparer\column-length.ps1"

<#

This function compares the column in an excel sheet to return back the names which occur the least. The function
could also be used to compare Contact Number, Email address etc but it can only compare values of only one type

Constraint: A name should be written in pattern <<First Name>> <<Last Name>>
            The sheet should contain only one type of value.
             
#>

function Compare-Excel {

    
    [CmdletBinding()] Param(
        [String] $str,
        [int] $num
    )

    process{
    
        write-host $str + "Hello"
        <#
            Open a excel sheet with input as parameter. 
        #>
        $excel=New-Object -ComObject Excel.application 
        $excel.Visible=$true
        $workbook=$excel.Workbooks
        $workbook=$workbook.Open($str)
        $worksheet=$workbook.Worksheets.item(1)

        <#
            Defining default HashMap
        #>

        $map = @{}

        <#
            Iterate the values present in coloumns of excel sheet and construct a HashMap 
        #>
        $intRow=1
        $intColumn=1

        $len = @{}
        <#
            Getting the number of columns
        #>
        $rows = $worksheet.UsedRange.Rows.Count
        $cols = $worksheet.UsedRange.Columns.Count

        for($i=0 ; $i -lt $cols ; $i++) {
            $k=[char]($i+65)
          $len[$i]=($excel.WorksheetFunction.CountIf($worksheet.Range($k+"1:"+$k+$rows), "<>") - 1)
        }

        $col_count=0
        do{ 
            write-host "Entered Here"
            for ( $i = 1 ; $i -le $len[$col_count] ; $i= $i+1){
                
                $value =  $worksheet.Cells.Item($i,$intColumn).Value()
                
                if(!$map.ContainsKey($value)){
                    $map.Add($value, 1)
                }
                else {
                    $map.$value = $map.$value + 1
                }
            }
            $intColumn=$intColumn+1
            $col_count++

        }while($worksheet.Cells.Item(1,$intColumn).Value() -ne $null)

        $i = $map.Count

        $map=$map.GetEnumerator() | Sort-Object value -Descending

        $map=$map.GetEnumerator() | Where-Object Value -LE $num

        $resultbook=$excel.Workbooks.Add()
        $resultsheet=$resultbook.Worksheets
        $resultsheet=$resultsheet.item(1)

        $i=1
        $map
        ForEach($item in $map.GetEnumerator()) { 
        $resultsheet.Cells.Item($i,1) = $item.Name
        $resultsheet.Cells.Item($i, 2) = $item.Value
        $i=$i + 1  }
                

    }

}