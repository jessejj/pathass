'this script makes the report the active window, un-Maximizes the window, sets the view to 'Web Layout', and sets the document to English

Set shell = CreateObject("Wscript.Shell")
    shell.AppActivate("Case")   'All PowerPath report windows start with 'Case' in the Microsoft Word title bar
Set shell = Nothing

Set w = GetObject(, "Word.Application")
Dim WindowState, ViewType
WindowState = w.WindowState 'whether or not the Window is Maximized, '1' being Maximized
ViewType = w.ActiveDocument.ActiveWindow.View.type  'how the window is currently displaced, e.g. Print View or Web Layout view

If Not WindowState = 0 Then w.WindowState = 0
If Not ViewType = 6 Then w.ActiveDocument.ActiveWindow.View.type = 6

w.ActiveDocument.Range.LanguageID = 4105  'wdEnglishCanadian

Set w = Nothing