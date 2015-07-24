Attribute VB_Name = "Pi_Helpers"
''this module contains the code related to working with PI

Sub Pi_CreateSampleData()

    'code will create a block of sampled data
    
    ''figure out how many tags there are
    Dim iTags As Integer
    Dim iRows As Integer
    
    'if there is 1,
    If Range("C1") = "" Then
        MsgBox "No tags found"
        Exit Sub
    ElseIf Range("D1") = "" Then
        iTags = 1
    Else
        iTags = RangeEnd(Range("C1"), xlToRight).count
    End If
        
    iRows = Range("A1")
    
    ''delete the formulas for whatever is in the sheet now
    Dim rngArr As Range
    Set rngArr = Intersect(RangeEnd(Range("B4"), xlDown, xlToRight), ActiveSheet.UsedRange)
    If Not rngArr Is Nothing Then
        rngArr.Formula = ""
    End If
    
    'LOGIC
    'for the first tag, create the formula for 2 columns
    'for the later tags, create formula in one column
    
    'formula for first column
    Range(Range("B4"), Range("B4").Offset(iRows, 1)).FormulaArray = _
        "=PISampDat(R1C3,R2C1,R2C2,R3C2,1,""CPAMHCC-PIMS01"")"
        
    Dim iCol As Integer
    For iCol = 1 To iTags - 1
    
        'for the first columnn, put the forumula down.  copy right for the others.
        If iCol = 1 Then
            Range(Range("C4").Offset(, iCol), Range("C4").Offset(iRows, iCol)).FormulaArray = _
                "=PISampDat(D$1,$A$2,$B$2,$B$3,0,""CPAMHCC-PIMS01"")"
        Else
            Range(Range("C4").Offset(, iCol), Range("C4").Offset(iRows, iCol)).FillRight
        End If
    
    Next iCol
    
    'format dates
    RangeEnd(Range("B4"), xlDown).NumberFormat = "mm/dd/yyyy HH:MM"

End Sub

Public Sub PiDataPull()

    Dim corner As Range
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    Set corner = ws.Range("A1")
    
    'Add in text
    ws.Range("A1").Formula = "=24*6"
    ws.Range("A2").Formula = "=B2-B1"
    ws.Range("B1") = 1
    ws.Range("B2") = Now
    ws.Range("A3") = "Interval"
    ws.Range("B3").Formula = "=(B2-A2)*24*60/A1&""m"""
    
    ws.Range("A2:B2").NumberFormat = "mm/dd/yyyy HH:MM"
    
    'get the pi tags
    Dim rngTags As Range
    Set rngTags = Range(corner, corner.End(xlToRight)).Offset(, 2).SpecialCells(xlCellTypeConstants)
    
    rngTags.Offset(1).FormulaR1C1 = "=PITagAtt(R[-1]C,""descriptor"")"
    rngTags.Offset(2).FormulaR1C1 = "=PITagAtt(R[-2]C,""engunits"")"

End Sub
