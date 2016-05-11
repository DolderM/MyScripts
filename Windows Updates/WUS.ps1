


(New-Object Net.WebClient).DownloadFile('https://gallery.technet.microsoft.com/scriptcenter/2d191bcd-3308-4edd-9de2-88dff796b0bc/file/41459/43/PSWindowsUpdate.zip',"$([environment]::getfolderpath('mydocuments'))\WindowsPowerShell\Modules\PSWindowsUpdate.zip")
(new-object -com shell.application).namespace("$([environment]::getfolderpath('mydocuments'))\WindowsPowerShell\Modules").CopyHere((new-object -com shell.application).namespace("$([environment]::getfolderpath('mydocuments'))\WindowsPowerShell\Modules\PSWindowsUpdate.zip").Items(),16)
Remove-Item "$([environment]::getfolderpath('mydocuments'))\WindowsPowerShell\Modules\PSWindowsUpdate.zip" -Force
Import-Module PSWindowsUpdate

Get-Command -Module PSWindowsUpdate