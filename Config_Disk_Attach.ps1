$VM = Get-VM "w7-base"
$ConfigVMDKPath =  "[prod_flash_03] Config/w8-base_1.vmdk"
New-HardDisk -VM $VM -DiskPath $ConfigVMDKPath -Persistence IndependentPersistent

$VM | Get-HardDisk | ? Filename -match $ConfigVMDKPath.split(" ")[1] | Remove-HardDisk -Confirm:$False