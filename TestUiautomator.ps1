Import-Module ".\bin\Debug\UIAutomator.dll"
Get-uiWindow -Name "*Calcu*" -Timeout 5 | Get-uiControl -Name Clear -Type button -Class button -AutoID clearButton | Invoke-uiSendKeys -Text 123456 -Click