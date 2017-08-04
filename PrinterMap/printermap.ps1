param(
    [string]$server,
    [string]$printershare
)
Try {
    if (Get-Printer | Where-Object {$_.ShareName -eq $sharename -and $_.DeviceType -eq "Print"}){
        return New-Object PSObject -Property @{'return'=0x4}
    } elseif (Get-Printer -ComputerName kyprint01 | Where-Object {$_.ShareName -eq $sharename -and $_.DeviceType -eq "Print"}) {
        Try {
            Add-Printer -ConnectionName $printershare | Out-Null
			return New-Object PSObject -Property @{'return'=0x0}
        } Catch {
            return New-Object PSObject -Property @{'return'=0x5}
        }
    } else {
        return New-Object PSObject -Property @{'return'=0x3}
    }
} Catch {
    return New-Object PSObject -Property @{'return'=0x2}
} Finally {
    Get-Job -Command "Add-Printer*" | Where {$_.State -eq "Completed"} | Remove-Job
}




