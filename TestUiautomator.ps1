Set-Location $PSScriptRoot
Remove-Module UIAutomator -Force -ErrorAction SilentlyContinue
Import-Module ".\bin\Debug\UIAutomator.dll"
#$win = Get-hWindow -Name "*Calcu*" -Timeout 5
#Get-uiWindow -Name "*Calcu*" -Timeout 5 | Get-uiControl -Name Clear -Type button -Class button -AutoID clearButton | Invoke-uiSendKeys -Text 123456 -Click
Get-uiWindow -Name "Hancom Office 2020" -Timeout 5000 | Get-uiControl -Name "OK(D)" -Class Button -Type Button -AutoID "1" -Timeout 10000