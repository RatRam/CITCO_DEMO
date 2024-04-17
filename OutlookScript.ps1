
if (Get-Process | Where-Object name -contains outlook) {
            Write-Output 'Outlook already running.'
            $outlook = New-Object -COM Outlook.Application;
            $mapi = $outlook.GetNameSpace("MAPI");
            $mapi.logon()
    }
else {
            Write-Output 'Starting Outlook application...'
		$exe = (Get-ItemProperty -Path Outlook.exe).'(default)'
            Get-Process Start-Process "Outlook" '(default)'
            Write-Output 'Outlook Reconnecting'
            $outlook = New-Object -COM Outlook.Application;
            $mapi = $outlook.GetNameSpace("MAPI");
            $mapi.logon()
    }

