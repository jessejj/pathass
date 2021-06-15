'this script makes the report the active window

Set shell = CreateObject("Wscript.Shell")
    shell.AppActivate("Case")   'All PowerPath report windows start with 'Case' in the Microsoft Word title bar
Set shell = Nothing
