function column-length{
    
    [cmdletbinding()]param(
        [paramater(Mandatory=$true, position=1)]
        $worksheet,
        [paramter(Mandatory=$true, position = 2)]
        [int]$column
    )
    process{
        
        $count = 0
        $row_num = 1

        while($worksheet.Cells.Item($row_num,$column).Value() -ne $null){
            $count = $count + 1
            $row_num = $row_num +1
        }

        return $count

    }
}