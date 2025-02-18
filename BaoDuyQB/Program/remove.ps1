$LangList = Get-WinUserLanguageList
$LangList[0].InputMethodTips.Remove("0409:00000409")
Set-WinUserLanguageList $LangList -Force