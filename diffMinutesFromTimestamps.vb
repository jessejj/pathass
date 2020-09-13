
Set ObjExcel = GetObject(, "Excel.Application")
Set objData = ObjExcel.Worksheets("Log")

Set ObjNSW = ObjExcel.Worksheets("NSW")

For blah = 1 To 100

    For FindLatest = 2 To 100 'find latest not recorded on NSW
        
        LatestRecord = objData.Cells(FindLatest, 1).Value
        NSWRecord = ObjNSW.Cells(2, 1).Value
                
        If LatestRecord <= NSWRecord Then
            Exit For
        End If
            
    Next
    
        If FindLatest <= 2 Then Exit For
    
        NewRecord = objData.Cells(FindLatest - 1, 1).Value
        
        MinDiff = (objData.Cells(FindLatest - 1, 1).Value - objData.Cells(FindLatest, 1).Value) * 24 * 60
        MinDiff = MinDiff - objData.Cells(FindLatest - 1, 3).Value
        'MsgBox MinDiff
        
        ObjNSW.Cells(2, 1).EntireRow.Insert
        ObjNSW.Cells(2, 1).Value = NewRecord
        
        If MinDiff < 480 Then
        
            If MinDiff > 30 Then
                'ObjNSW.Cells(2, 2).Value = objData.Cells(FindLatest - 1, 2).Value
                ObjNSW.Cells(2, 2).Value = 30
                
                ObjNSW.Cells(2, 1).EntireRow.Insert
                ObjNSW.Cells(2, 1).Value = NewRecord
                'ObjNSW.Cells(2, 2).Value = objData.Cells(FindLatest - 1, 2).Value
                'ObjNSW.Cells(2, 3).Value = objData.Cells(FindLatest - 1, 3).Value
                ObjNSW.Cells(2, 2).Value = MinDiff - 30
                
            Else
            
               ' ObjNSW.Cells(2, 2).Value = objData.Cells(FindLatest - 1, 2).Value
               ' ObjNSW.Cells(2, 3).Value = objData.Cells(FindLatest - 1, 3).Value
                ObjNSW.Cells(2, 2).Value = MinDiff
                
            End If
            
        End If
        
        
        
        
        'ObjNSW.Cells(2, 1).EntireRow.Insert
        'ObjNSW.Cells(2, 1).Value = NewRecord
        'ObjNSW.Cells(2, 2).Value = objData.Cells(FindLatest - 1, 2).Value
        
        'MinDiff = (objData.Cells(FindLatest - 1, 1).Value - objData.Cells(FindLatest, 1).Value) * 24 * 60
        
        'ObjNSW.Cells(2, 3).Value = MinDiff

Next

Exit Sub



For i = 1 To 1000
NumberOfRecordsChecked = NumberOfRecordsChecked + 1

    If ObjNSW.Cells(NSWLastRow, 2).Value = LastRowTime And ObjNSW.Cells(NSWLastRow, 1).Value = LastRowDate Then
        ValuePresent = True
        Exit For
    Else
        If DataLastRow = 1 Then Exit For
        DataLastRow = objData.Cells(DataLastRow, 1).Offset(-1, 0).Row
        LastRowDate = objData.Cells(DataLastRow, DateColumn).Value
        LastRowTime = objData.Cells(DataLastRow, TimeColumn).Value
    End If
   ' If ValuePresent = True Then
    '    DataLastRow = objData.Cells(DataLastRow, 1).Offset(-1, 0).Row
     '   LastRowDate = objData.Cells(DataLastRow, DateColumn).Value
    '    LastRowTime = objData.Cells(DataLastRow, TimeColumn).Value
    'End If
'If i = 100 Then
'Answer = MsgBox("The amount of records not present exceeds 100. Are you sure you want to continue calculations?", vbYesNo)
'    If Answer = vbYes Then i = 1 Else Exit Sub
'End If
Next

'now start doing differences
For i = 1 To 1000
    If objData.Cells(DataLastRow, DateColumn).Value = "" Then Exit For
    If DataLastRow = 1 Then
        DataLastRow = 2
        LastRowDate = objData.Cells(DataLastRow, DateColumn).Value
        LastRowTime = objData.Cells(DataLastRow, TimeColumn).Value
    End If
 
    
    TimeDifference = Val(objData.Cells(DataLastRow, TimeColumn).Value) - Val(objData.Cells(DataLastRow - 1, TimeColumn).Value)
    'TimeDifference = Round(TimeDifference * 24 * 60, 0)    'prefer round down bc Word rounds down report min
    
    'MsgBox TimeDifference
    TimeDifference = Int(TimeDifference * 24 * 60) 'Int function rounds down
    
    ReportEditingTime = Val(objData.Cells(DataLastRow, ReportEditingTimeColumn).Value)
    
    If TimeDifference >= ReportEditingTime Then TimeDifference = TimeDifference - ReportEditingTime
    
    If TimeDifference > 30 Then 'trim for breaks
        TrimValue = TimeDifference - 30
        ObjNSW.Cells(NSWLastRow, DateColumn).Value = LastRowDate
        ObjNSW.Cells(NSWLastRow, TimeColumn).Value = LastRowTime
        ObjNSW.Cells(NSWLastRow, NSWMinutesColumn).Value = TrimValue
        TimeDifference = 30
        NSWLastRow = ObjNSW.Cells(NSWLastRow, 1).Offset(1, 0).Row
    End If
    
    If TimeDifference >= 1 Then

    ObjNSW.Cells(NSWLastRow, DateColumn).Value = LastRowDate
    ObjNSW.Cells(NSWLastRow, TimeColumn).Value = LastRowTime
    ObjNSW.Cells(NSWLastRow, NSWMinutesColumn).Value = TimeDifference
    ObjNSW.Cells(NSWLastRow, NSWPrimarySpecimenColumn).Value = objData.Cells(DataLastRow, PrimarySpecimenColumn).Value
    ObjNSW.Cells(NSWLastRow, NSWAccessionNumberColumn).Value = objData.Cells(DataLastRow, AccessionNumberColumn).Value
    
    
    NSWLastRow = ObjNSW.Cells(NSWLastRow, 1).Offset(1, 0).Row
    End If
    
    DataLastRow = objData.Cells(DataLastRow, 1).Offset(1, 0).Row
    LastRowDate = objData.Cells(DataLastRow, DateColumn).Value
    LastRowTime = objData.Cells(DataLastRow, TimeColumn).Value
    

DateCheck = objData.Cells(DataLastRow, DateColumn).Value

If i = 1000 Then
Answer = MsgBox("The number of calculations exceeds 100. Are you sure you want to continue calculations?", vbYesNo)
    If Answer = vbYes Then i = 1 Else Exit Sub
End If
Next




'If ValuePresent <> True Then
 '   objNSW.Cells(NSWLastRow, DateColumn).Value = LastRowDate
 '   objNSW.Cells(NSWLastRow, TimeColumn).Value = LastRowTime
'End If

Set ObjExcel = Nothing