

$PoolList = @(
    "w7-o7-ie9",
    "w7-o7-ie10",
    "w7-o7-ie11",
    "w7-o10-ie9",
    "w7-o10-ie10",
    "w7-o10-ie11",
    "w7-o13-ie9",
    "w7-o13-ie10",
    "w7-o13-ie11",
    "wV-o7-ie8",
    "wV-o7-ie9",
    "wV-o10-ie8",
    "wV-o10-ie9",
    "w8-o7-ie11",
    "w8-o10-ie11",
    "w8-o13-ie11"
    )

$PoolSettings = @{
    maximumCount            = "3"
    minimumCount            = "2"
    headroomCount           = "1"
    minprovisioneddesktops  = "0"
    autoLogoffTime          = "15"
    refreshPolicyType       = "Never"
    deletePolicy            = "Default"
    powerPolicy             = "PowerOff"
    }

Add-PSSnapin VMware.View.Broker 
Add-PSSnapin VMware.VimAutomation.Core
$ErrorActionPreference = "stop"

$VcenterServer = "View.contoso.corp"
$GuestPassword = "password"
$GuestUsername = "administrator"
$ConfigVMDKPath =  "[datastore] Folder/ConfigDisk.vmdk"
$Cluster = "ViewCluster"
$OU = "OU=View"
$Datastore = "datastore"
$EntitlementGroup = "Domain Users"

# This script redefines the variables above to work with my environment.  
# It is not included for obvious reasons, but matches the above block.
. .\MyEnvironment.ps1


$GuestCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $GuestUsername,($GuestPassword | ConvertTo-SecureString -AsPlainText -Force)
Connect-VIServer $vCenterServer 
$ComposerDomain = Get-ComposerDomain
$vc_id = $ComposerDomain.vc_id
$Composer_AD_ID = $ComposerDomain.composer_ad_id
$EntitlementSID = Get-ADGroup $EntitlementGroup | Select-Object -ExpandProperty SID | Select-Object -ExpandProperty Value
$ProgressPreference = "SilentlyContinue"

#### Helper Functions ###

Function Parse-Poolname ($PoolName) {
    $Versions = $PoolName.split("-")
    $PoolObject = @{OSVer = $Versions[0];OfficeVer = $Versions[1]; IEVer = $Versions[2]}
    return $PoolObject
    }

Function Add-ConfigVMDK($VM){
    Write-Host "Attaching setup VMDK: $ConfigVMDKPath"
    New-HardDisk -VM $VM -DiskPath $ConfigVMDKPath -Persistence IndependentPersistent | Out-Null
    }

Function Remove-ConfigVMDK($VM) {
    Write-Host "Removing setup VMDK: $ConfigVMDKPath"
    $VM | Get-HardDisk | ? Filename -match $ConfigVMDKPath.split(" ")[1] | Remove-HardDisk -Confirm:$False
    }

Function Check-VMReadiness($VM, $Type) {
    switch ($Type)
        {
        "PoweredOn" {
            do {"Zzz...";Start-Sleep -Seconds 5}
            until ((Get-VM $BaseVM).PowerState -match "PoweredOn")}
        "PoweredOff" {
            do {"Zzz...";Start-Sleep -Seconds 5}
            until ((Get-VM $BaseVM).PowerState -match "PoweredOff")}
        "Script" {
            Write-Host "Waiting for $OSVer-base to get ready..."
            do {"Zzz...";Start-Sleep -Seconds 5}
            while (!(Invoke-VMScript -VM $BaseVM -ScriptText "ECHO TEST" -GuestCredential $GuestCredential -ErrorAction Ignore ))}
        }
    }

### Snapshot Management Function ###

Function Make-Snapshot($PoolName){
    $PoolNameObject = Parse-Poolname $PoolName
    $OSVer = $PoolNameObject.OSVer
    $IEVer = $PoolNameObject.IEVer
    $OfficeVer = $PoolNameObject.OfficeVer

    $OfficeBase = "$OSVer-$OfficeVer-base"
    $OSBase = "$OSVer-base"
    $BaseVM = Get-VM $OSBase


    

    if (-NOT ($BaseVM | Get-Snapshot | ? Name -match $OfficeBase) -AND ($IEVer -notmatch "base")){
        Write-Host "Building Office base..."
        Make-Snapshot $OfficeBase
        } 

    if (($BaseVM | Get-Snapshot | ? Name -match $OfficeBase) -AND ($IEVer -notmatch "base")){
        Write-Host "Reverting to base snapshot: $OfficeBase"
        $OfficeSnapshot = $BaseVM | Get-Snapshot | ? Name -match $OfficeBase
        Set-VM -VM $BaseVM -Snapshot $OfficeSnapshot.Name -Confirm:$false | Out-Null
        }

    if (($IEVer -match "base")) {
        Write-Host "Reverting to base snapshot: $OSBase"
        Set-VM -VM $BaseVM -SnapShot $OSBase -Confirm:$false | Out-Null
        }
    
   
    
    Add-ConfigVMDK $BaseVM
    Write-Host "Starting $OSVer-base..."
    Start-VM -VM $BaseVM | Out-Null
    Start-Sleep -Seconds 5
    Check-VMReadiness $BaseVM "Script"

    if ($IEVer -match "base") {
        Write-Host "Installing Office: $OfficeVer"
        Invoke-VMScript -ScriptText "start-process -filepath `"z:\setup\$OfficeVer.cmd`" -wait"   -VM $BaseVM -GuestCredential $GuestCredential
        
        Write-Host "Running Chocolatey Installation Script"
        Invoke-VMScript -ScriptText "start-process -filepath `"z:\ChocoInstalls.cmd`" -wait"   -VM $BaseVM -GuestCredential $GuestCredential
        }

    if ($IEVer -notmatch "base") {
        Write-Host "Installing IE: $IEVer" 
        Invoke-VMScript -ScriptText "start-process -filepath `"z:\setup\$IEVer.cmd`" -wait"   -VM $BaseVM -GuestCredential $GuestCredential

        Write-Host "Restarting $OSBase..."
        Restart-VMGuest -VM $BaseVM | Out-Null
        Start-Sleep -Seconds 30
        write-host "Waiting for $OSBase to boot up..."
        Check-VMReadiness $BaseVM "PoweredOn"
        write-host "Waiting for $OSBase to finish installing IE..."
        Check-VMReadiness $BaseVM "Script"
        }


    Write-Host "Shutting down $OSBase..."
    Shutdown-VMGuest -VM $BaseVM -Confirm:$false | Out-Null
    Check-VMReadiness $BaseVM "PoweredOff"
    Remove-ConfigVMDK $BaseVM
    if (($BaseVM | Get-Snapshot).name -Match $PoolName) {
        "Removing old snapshot: $(($BaseVM | Get-Snapshot | ? Name -Match $PoolName).Name)"; 
        $BaseVM | Get-Snapshot | ? Name -match $PoolName | Remove-Snapshot -confirm:$False
        }
    $NewSnapshotName = "$PoolName-$(get-date -format yyMMddHHmm)"
    write-host "Making new snapshot: $NewSnapshotName"
    New-Snapshot -VM $BaseVM -Name $NewSnapshotName | Out-Null
    Start-Sleep -Seconds 30
    write-host "Reverting to base snapshot..."
    Set-VM -VM $BaseVM -SnapShot $OSBase -Confirm:$false | Out-Null
    }

### Pool Management Functions ###

Function Make-Pool ($PoolName) {
    $PoolNameObject = Parse-Poolname $PoolName
    $OSVer = $PoolNameObject.OSVer
    $IEVer = $PoolNameObject.IEVer
    $OfficeVer = $PoolNameObject.OfficeVer
    $TemplateName = "$OSVer-base"
    $BaseVM = Get-VM $templateName

    $Snapshot = $BaseVM | Get-Snapshot | ? Name -match $PoolName | Select-Object -first 1
    If (!($Snapshot)) {Throw "No snapshot matching $PoolName on $TemplateName"}
    $SnapshotName = $Snapshot.name
    $SnapshotPath = "/$TemplateName/$($Snapshot.Parent)/$SnapshotName"
    


    $LinkedClonePoolSettings = @{
        pool_id                 = $PoolName
        description             = $PoolName
        displayName             = $PoolName
        vc_id                   = $vc_id
        persistence             = "NonPersistent"
        resourcePoolPath        = "/$Cluster/host/View/Resources"
        vmFolderPath            = "/$Cluster/vm/View"
        parentVMPath            = "/$Cluster/vm/View_Templates/$TemplateName"
        parentSnapshotPath      = $SnapshotPath
        datastoreSpecs          = "[Conservative,OS,data]/$Cluster/host/View/$Datastore"
        composer_ad_id          = $Composer_AD_ID
        organizationalUnit      = $OU
        namePrefix              = "$PoolName-{n}"
        IsProvisioningEnabled   = $True
        SuspendProvisioningOnError = $True
        IsUserResetAllowed      = $True
        DefaultProtocol         = "PCOIP"
        allowProtocolOverride   = $True
        }
    $LinkedClonePoolSettings += $PoolSettings

    Get-ViewVC -ServerName $VcenterServer | Out-Null
    Add-AutomaticLinkedClonePool @LinkedClonePoolSettings
    Add-PoolEntitlement -Pool_id $LinkedClonePoolSettings.pool_id -Sid $EntitlementSID
    Start-Sleep -Seconds 10
    $poolObject = New-Object System.DirectoryServices.DirectoryEntry "LDAP://localhost:389/cn=$poolName,ou=Applications,dc=vdi,dc=vmware,dc=int"
    $poolObject.'pae-ServerProtocolLevel' = @('BLAST','PCOIP','RDP')
    $poolObject.CommitChanges()

    }

Function Recompose-Pool ($PoolName) {
    $PoolNameObject = Parse-Poolname $PoolName
    $OSVer = $PoolNameObject.OSVer
    $IEVer = $PoolNameObject.IEVer
    $OfficeVer = $PoolNameObject.OfficeVer
    $OSBase = "$OSVer-base"
    $templateName = $OSBase
    $MachineID = Get-DesktopVM | ? Pool_ID -Match $PoolName | select -first 1 | select -expandproperty machine_ID

    $Snapshot = Get-VM $templateName | Get-Snapshot | ? Name -match $PoolName | Select-Object -first 1
    If (!($Snapshot)) {Throw "No snapshot matching $PoolName on $TemplateName"}
    $SnapshotName = $Snapshot.name
    $SnapshotPath = "/$templateName/$($Snapshot.Parent)/$SnapshotName"
    
    $SendLinkedCloneRecompose = @{
        ParentVMPath    =  "/$Cluster/vm/View_Templates/$templateName"
        parentSnapshotPath      = $SnapshotPath
        Machine_ID      = $MachineID
        ForceLogoff     = $True
        StopOnError     = $False
        Schedule        = (Get-Date)
        }
     Send-LinkedCloneRecompose @SendLinkedCloneRecompose
    }


### Convenience Functions ###


Function Create-OSPools ($OSVersion){
    $PoolNames = $PoolList -match "$OSVersion"
    $PoolNames | % {Make-Pool $_}
    }

Function Create-AllPools {
    $PoolList | % {Make-Pool $_}
    }

Function Recompose-OSPools ($OSVer) {
    $Pools = Get-Pool | ? Pool_ID -Match $OSVer 
    $Pools | % {Write-Host "Recomposing pool: $($_.pool_id)"; Recompose-Pool $_.pool_id}
    }

Function Recompose-AllPools {
    $Pools = Get-Pool
    $Pools | % {Recompose-Pool $_.pool_id}
    }

Function Regenerate-OSSnapshots ($OSVer, $RecomposePools = $False){
    ClearOfficeBase $OSVer
    $PoolNames = $PoolList -match "$OSVer"
    $PoolNames | % {write-host "Making snapshot: $_"; Make-Snapshot $_}
    If ($RecomposePools) {Recompose-OSPools $OSVer}
    }

Function Regenerate-AllSnapshots ($RecomposePools = $False){
    ClearOfficeBase
    $PoolList | % {Write-Host "Making snapshot: $_"; Make-Snapshot $_}
    If ($RecomposePools) {Recompose-AllPools}
    }

Function ClearOfficeBase($OSVer){
    if ($OSVer) {
        $BaseVM = Get-VM "$OSver-base" 
        $BaseVM | Get-Snapshot | ? Name -match "$OSVer-o[0-9]+-base" | Remove-Snapshot -Confirm:$False | Out-Null
    } else {
        $BaseVMs = Get-VM | ? Name -match "w.-base"
        foreach ($BaseVM in $BaseVMs) {
            $Snapshots = $BaseVM | Get-Snapshot | ? Name -match "w[v78]-o[0-9]+-base"
            Write-Host "Removing snapshots:"
            $Snapshots | Select -expand Name
            $Snapshots | Remove-Snapshot -Confirm:$False | Out-Null
            }
        }
    }


