#Script to remove songs from sub folders and copy them in one single folder.

function extract-songs{

    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory=$True,Position=1)]
            [string]$inputpath,
        [Parameter(
            Mandatory=$True,Position=2)]
            [string]$outputpath
    )
    
    process{
        
        #Checking whether inputpath and outputpath exist or not
        if(!(Test-Path $inputpath)){
           #$inputpath_exist=$false
           return 'No path Exist. Kindly Check the path' 
        }

        #If output path is not present. It creates the output path
        if(!(Test-Path $outputpath)){
            New-Item $outputpath -ItemType Directory -force
        }

        #Getting the list of all files and folder in Song Folder.
        $dir = Get-ChildItem $inputpath -Recurse

        #Fetching the list of only .mp3 files from the folder. 
        $list = $dir | where { $_.Extension -eq '.mp3' }

        # '$count' gives the number of files in the folder
        $count=$list.Count

        #Copying the files into the output folder
        Foreach($a in $list){
        $tmp=$a.Name
        Copy-item -path $a.FullName -Destination $outputpath\$tmp
        }
        
    }

}