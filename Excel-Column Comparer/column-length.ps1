Function column-length{
    
    [cmdletbinding()]param(
     
        [parameter(Mandatory=$true, position=1)]
        $worksheet,
     
        [parameter(Mandatory=$true, position = 2)]
        [int]$column
    
    )
    process{
        
        $count = 0
        $row_num = 1

        while($worksheet.Cells.Item($row_num,$column).Value() -ne $null){
            $count = $count + 1
            $row_num = $row_num +1
            write-host $count
        }

        return $count

    }
}