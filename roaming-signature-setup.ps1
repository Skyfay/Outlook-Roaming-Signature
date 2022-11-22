# creator: Skyfay
# Support: support@skyfay.ch
# Last edit: 21.11.2022

#‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾#
#->                                                                       Funcions                                                              <-#
#_________________________________________________________________________________________________________________________________________________#

function disable_roaming_signature {
    $path1 = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Setup'
    New-ItemProperty -Path $path1 -Name 'DisableRoamingSignaturesTemporaryToggle' -Value 1 -PropertyType DWord

}

function enable_roaming_signature {
    $path1 = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Setup'
    Remove-ItemProperty -Path $path1 -Name 'DisableRoamingSignaturesTemporaryToggle'
}

function status_check {
    $error.clear() 
    try {
        Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Setup' -Name 'DisableRoamingSignaturesTemporaryToggle' -ErrorAction Stop | out-null
    }
    catch {
        Write-Host "Status: Outlook Signature Roaming is currently active." -ForegroundColor green
    }
    if (!$error) {
        Write-Host "Status: Outlook Signature Roaming is currently disabled." -ForegroundColor red
    }
}

#‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾#
#->                                                                     Main Code                                                               <-#
#_________________________________________________________________________________________________________________________________________________#

cls
Write-Host "With this script you can disable and enable roaming signatures for Outlook."
Start-Sleep 1
Write-Host `n
status_check
Write-Host `n
Write-Host "(1)  -  disable roaming signatures"
Write-Host "(2)  -  enable roaming signatures"
Write-Host `n

$value = Read-Host "Do you want to enable or disable roaming signatures?"


switch ($value) {
    1 {disable_roaming_signature}
    2 {enable_roaming_signature}
    Default {
    cls
    Write-Host "Please enter a value between 1-2! $value is invalide."
    Start-Sleep 4
    Exit
    }
}

cls

Write-Host "You have successfully reconfigured roaming signatures for Outlook"
Start-Sleep 2
