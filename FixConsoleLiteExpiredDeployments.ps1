

Function Get-CollectionsInFolder {
    [CmdletBinding()] 
    PARAM ( [Parameter(Position=1)] $SiteServer,
            [Parameter(Position=2)] $SiteCode,
            [Parameter(Position=3)] $FolderNodeID  )  

    $CollectionObjects = @()
    $CollectionObjects = Get-WmiObject -ComputerName $SiteServer  -Namespace "Root\SMS\Site_$SiteCode" -Query "SELECT * FROM SMS_ObjectContainerItem WHERE ContainerNodeID = $FolderNodeID"
    Return $CollectionObjects
}


Function Get-DeploymentForCollection  {
    [CmdletBinding()] 
    PARAM ( [Parameter(Position=1)] $SiteServer,
            [Parameter(Position=2)] $SiteCode,
            [Parameter(Position=3)] $CollectionID  ) 
    
    $DeploymentObjects = @()
    $DeploymentObjects = Get-WmiObject -ComputerName $SiteServer  -Namespace "Root\SMS\Site_$SiteCode" -Query "SELECT * FROM SMS_Advertisement WHERE CollectionID = '$CollectionID'"
    Return $DeploymentObjects
}




############################
#   MAIN
############################

$FolderNodeID_ConsoleLiteAutomated = 0xc4       #PROD
$SiteServer = "XSPW10W200P"                     #PROD
$SiteCode = "P00"                               #PROD


$LogFile = "C:\Temp\FixConsoleLiteExpirdUpdates.log"
$TodaysDate = Get-Date
$Collections = @()
$Deployments = @()
        ("Fixing Expired ConsoleLite deployments "+$TodaysDate) | Out-File $LogFile   -Append

$Collections = Get-CollectionsInFolder -SiteServer $SiteServer -SiteCode $SiteCode -FolderNodeID $FolderNodeID_ConsoleLiteAutomated

ForEach ( $Collection in $Collections ) {
   $Deployments += Get-DeploymentForCollection -SiteServer $SiteServer -SiteCode $SiteCode -CollectionID $Collection.InstanceKey
}

ForEach ( $Deployment in $Deployments ) {
    $WMIDeployment = [wmi]$Deployment.__PATH
    If ( $WMIDeployment.ExpirationTimeEnabled  -eq $True ) {
        ("CollectionID: "+$WMIDeployment.CollectionID+"     AdvertisementID: "+$WMIDeployment.AdvertisementID+"     ProgramName: "+$WMIDeployment.ProgramName+"     AdvertisementName: "+$WMIDeployment.AdvertisementName+"     ExpirationTimeEnabled: "+$WMIDeployment.ExpirationTimeEnabled ) | Out-File $LogFile   -Append
        $WMIDeployment.ExpirationTimeEnabled = $False
        $Result = $WMIDeployment.Put()
    }

}   