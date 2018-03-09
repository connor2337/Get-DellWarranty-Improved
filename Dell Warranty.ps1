#----------------------------------------------------------------------------------------------------------------------------#
#                                                                                                                            #
# The core of this script (about 40%) was written by Natascia Heil and published in the TechNet Gallery under MIT licensing  #
# The rest was written by me (Connor Carroll)                                                                                #
#                                                                                                                            #
# Please carefully look through through the code as there are things you must change for this script to work properly        #
#                                                                                                                            #
#----------------------------------------------------------------------------------------------------------------------------#
#                                                                                                                            #
# USAGE:                                                                                                                     #
#                                                                                                                            #
# When you run this script, you can either use the first inputbox to point to a path with a newline delimited txt file.      #
# This first option will create a CSV in the same directory that you ran the script in.                                      #
#                                                                                                                            #
# OR                                                                                                                         #
#                                                                                                                            #
# You can leave the first inputbox empty, click okay, and type in one hostname in the next inputbox, which will show the     #
# warranty information for that specific PC in a console window.                                                             #
#                                                                                                                            #
#----------------------------------------------------------------------------------------------------------------------------#

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

Function Get-DellWarranty 
{ 
    # Define Parameters
    [CmdletBinding()] 
    Param(   
        [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true, Mandatory=$true)]      
        [Alias("Serial","SerialNumber")] 
        [String]$ServiceTag
         
        , 
         [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true, Mandatory=$false)]      
        [String]$macdesc

        , 
         [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true, Mandatory=$false)]      
        [String]$macaddr
        ,
         
        [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true, Mandatory=$false)]      
        [String]$curUser
        ,
        
        [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true, Mandatory=$false)]      
        [String]$hostname
        , 
          
        [Parameter(Mandatory=$false)]   
        [String]$ApiKey = 'APIKEY123'      <# PUT YOUR API KEY HERE AS STRING #>
        , 
          
        [Parameter(Mandatory=$false)]   
        [Switch]$Dev = $true
        , 
 
        [Parameter(Mandatory=$False)] 
        [INT]$TagLimit = 25 
 
    )
 
    Begin # Setup variables functions 

    {    
    
        
        If (($Dev))  
        {  
            $Server = "https://sandbox.api.dell.com/support/assetinfo/v4/getassetwarranty/"       
        }  
        else  
        {  
            $Server = "https://api.dell.com/support/assetinfo/v4/getassetwarranty/"  
        } 
 
        Function Submit-Tag
        { 
           Param( 
           [String]$Tag
           , 
           [URI]$URL 
           ,
           [String]$macdesc
           ,
           [String]$macaddr
           ,
           [String]$curUser
           ,
           [String]$hostname
           ) 
                    
            Try  
            { 
                $Warranty = Invoke-RestMethod -URI $URL -Method GET -ContentType 'Application/xml' 
            } 
            Catch 
            { 
                Write-Error $Error[0] 
                Break 
            } 
     
            $Global:Get_DellWarrantyXML = $Warranty  
     
            foreach ($Asset in $Warranty.AssetWarrantyDTO.AssetWarrantyResponse.AssetWarrantyResponse) 
            { 
                Foreach ($Entitlement in $Asset.assetentitlementdata.assetentitlement | Where-Object ServiceLevelDescription -NE 'Dell Digitial Delivery'`
                 | Where-Object EntitlementType -EQ 'INITIAL') 


                { 
                    $row = New-Object PSObject 
                    $row | Add-Member -Name "Serial" -MemberType NoteProperty -Value $Asset.assetheaderdata.ServiceTag 
                    $row | Add-Member -Name "Model" -MemberType NoteProperty -Value $Asset.productheaderdata.SystemDescription
                    $row | Add-Member -Name "MAC Description" -MemberType NoteProperty -Value $macdesc
                    $row | Add-Member -Name "MAC Address" -MemberType NoteProperty -Value $macaddr
                    $row | Add-Member -Name "Current User" -MemberType NoteProperty -Value $curUser
                    $row | Add-Member -Name "Hostname" -MemberType NoteProperty -Value $hostname

                    if ($entitlement.ServiceLevelDescription.nil)# Look for Nulls in the XML 
                        {$row | Add-Member -Name "ServiceLevelDescription" -MemberType NoteProperty -Value $NULL} 
                        else 
                        {$row | Add-Member -Name "ServiceLevelDescription" -MemberType NoteProperty -Value $entitlement.ServiceLevelDescription} 
 

                    $row | Add-Member -Name "StartDate" -MemberType NoteProperty -Value $entitlement.StartDate
                    $row | Add-Member -Name "EndDate" -MemberType NoteProperty -Value $entitlement.EndDate

                    return $row
                } 
            } 
        }# Push tags to dell 
 
 
        $URI = $Server + $ServiceTag + "?apikey=" + $Apikey
        $row = Submit-Tag -Tag $ServiceTag -URL $URI -macdesc $macdesc -macaddr $macaddr -curUser $curUser -hostname $hostname
        return $row
    
    } 
}

$txtpath=[Microsoft.VisualBasic.Interaction]::InputBox("Enter a path to a hostname list:", "Dell Warranty Info")
$inputbox=[Microsoft.VisualBasic.Interaction]::InputBox("Enter a Hostname:", "Dell Warranty Info")

if($inputbox -eq '')
{
    $hostinput=Get-Content -path $txtpath
    $hosts = $hostinput.Split([Environment]::NewLine)
    
    $Script:Table = New-Object System.Collections.ArrayList
    foreach($hosty in $hosts)
    { 
       Try{
       $macaddr =   Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $hosty -ErrorAction Stop | `
                    Where-Object DNSDomain -EQ <#TESTY.MCTESTERSON.LOCAL#> | Select Description, MACAddress <#---------------- Enter DNSDomainName to ensure it grabs the right MAC #>
       $macaddr = $macaddr | Select Description, MACAddress       
       $svcDell = Get-WmiObject Win32_BIOS -ComputerName $hosty | %{$_.SerialNumber}       
       if ($svcDell.Length -eq 7){
            $curUser = Get-WmiObject –ComputerName $hosty –Class Win32_ComputerSystem | Select-Object UserName
            $hostname = $hosty
            $row = Get-DellWarranty $svcDell $macaddr.Description $macaddr.MACAddress $curUser.username $hostname
            $Table.add($row)
              
       }
       }Catch{
       }
    }
    $Date = (Get-Date -UFormat "%m - %d - %Y")
    $FilePath = ".\$Date Dell Warranty Report.csv"
    $Table | Export-Csv $FilePath -NoTypeInformation

}

else

{
    $hosty=$inputbox
    $macaddr =  Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $hosty | `
                Where-Object DNSDomain -EQ <#TESTY.MCTESTERSON.LOCAL#> | Select Description, MACAddress <#---------------- Enter DNSDomainName to ensure it grabs the right MAC #>
    $macaddr = $macaddr | Select Description, MACAddress
    $svcDell = Get-WmiObject Win32_BIOS -ComputerName $hosty | %{$_.SerialNumber}
    $curUser = Get-WmiObject –ComputerName $hosty –Class Win32_ComputerSystem | Select-Object UserName
    $hostname = $hosty
    Get-DellWarranty $svcDell $macaddr.Description $macaddr.MACAddress $curUser.username $hostname | Format-Table
    pause
}
