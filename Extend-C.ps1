function Get-ObjectList
{
    param($key1)
    
    $id= whoami
    $trimmed_id = $id.toString().split('\')[1]

    if($key1 -eq "computer"){$content = Get-Content C:\Users\$trimmed_id\Desktop\computers.txt|Where-Object{$_.trim() -ne ''}}
    elseif($key1 -eq "group"){$content = Get-Content C:\Users\$trimmed_id\Desktop\groups.txt|Where-Object{$_.trim() -ne ''}}
    elseif($key1 -eq "LANID"){$content = Get-Content C:\Users\$trimmed_id\Desktop\lanIDs.txt|Where-Object{$_.trim() -ne ''}}
    elseif($key1 -eq "country"){$content = Get-Content C:\Users\$trimmed_id\Desktop\countries.txt|Where-Object{$_.trim() -ne ''}}

    return $content

}

Function Connect-DCOM
{
    param($comp)

    $SessionOption = New-CimSessionOption -Protocol DCOM

    try{
        
        New-CimSession -ComputerName $comp -SessionOption $SessionOption -ErrorAction Stop |Out-Null

        return "Success"
    }

    catch{

        return "Failed"
    }

    
}

Function Update-Storage
{
    param($comp)
    
    try{
        
        Update-HostStorageCache -CimSession (Get-CimSession|Where-Object{$_.ComputerName -eq $comp}) -ErrorAction Stop

        return "Success"
    }
    
    catch{

        return "Failed"
    }
}

function Get-PartitionSizeProperties
{
    param($comp,$drive_letter)

    $size_prop = (Get-PartitionSupportedSize -CimSession (Get-CimSession|Where-Object{$_.ComputerName -eq $comp}) -DriveLetter $drive_letter)

    return $size_prop
}

function Extend-Drive
{
    param($comp,$drive_letter,$size_prop)

    try{
        Resize-Partition -DriveLetter $drive_letter -Size $size_prop.SizeMax -ErrorAction Stop
        
        return "Success"
    }

    catch{
        return "Failed"
    }
}


function Get-Hash
{
    param($code,$computer)

    if($code -eq 1)
    {
          $hash=@{
                       "Computer" = $computer
                       "DCOM" = "Success"
                       "Scanning"  = "Success"
                       "C Extend" = "Success"
                       
                 }

          return $hash
    }

    elseif($code -eq 2)
    {
          $hash=@{
                       "Computer" = $computer
                       "DCOM" = "Failed"
                       "Scanning"  = "Failed"
                       "C Extend" = "Failed"
                       
                 }

          return $hash
    }

     elseif($code -eq 3)
    {
          $hash=@{
                       "Computer" = $computer
                       "DCOM" = "Success"
                       "Scanning"  = "Failed"
                       "C Extend" = "Failed"
                       
                 }

          return $hash
    }

     elseif($code -eq 4)
    {
          $hash=@{
                       "Computer" = $computer
                       "DCOM" = "Success"
                       "Scanning"  = "Success"
                       "C Extend" = "Failed"
                       
                 }

          return $hash
    }
}

$ext_C =@()


foreach($comp in (Get-ObjectList computer))
{
    
        Write-Host -NoNewline "Extending C drive in $comp..."
        
        $DCOM_result = Connect-DCOM $comp

        if($DCOM_result -eq "Success")
        {
            #Scan the storage
            $update_storage_result = Update-Storage $comp

            if($update_storage_result -eq "Success")
            {
                #Get C drive size properties
                $size_prop = Get-PartitionSizeProperties $comp 'C'

                #Extend the C drive
                #$extend_result =  Extend-Drive $comp 'C' $size_prop

                try{
                     Resize-Partition -CimSession (Get-CimSession|Where-Object{$_.ComputerName -eq $comp}) -DriveLetter 'C' -Size $size_prop.SizeMax
                     
                     $ext_C += New-Object psobject -Property (Get-Hash 1 $comp)
                }

                catch{

                    $ext_C += New-Object psobject -Property (Get-Hash 4 $comp)
                }



                
            }

            
            
            else{$ext_C += New-Object psobject -Property (Get-Hash 3 $comp)}
            
            
        }

        else
        {
            $ext_C += New-Object psobject -Property (Get-Hash 2 $comp)
        }
    
        Write-Host "Done"
}
        
        $ext_C|Select-Object Computer,DCOM,Scanning,'C Extend'



$stamp = get-date -Format "dd-MMM-yyyy HHmm"
$my_name = whoami | Out-String
$user_name = $my_name.Split("\")[1].Trim()
$current_directory = pwd

        $ext_C|Select-Object Computer,DCOM,Scanning,'C Extend'|Export-Csv C_Ext_$stamp"_"$user_name.csv -NoTypeInformation

        Get-CimSession |Remove-CimSession