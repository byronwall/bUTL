Attribute VB_Name = "Pi_Helpers"
''this module contains the code related to working with PI

Sub Pi_CreateSampleData()

    'code will create a block of sampled data
    
    ''figure out how many tags there are
    Dim int_tags As Integer
    Dim int_rows As Integer
    
    'if there is 1,
    If Range("C1") = "" Then
        MsgBox "No tags found"
        Exit Sub
    ElseIf Range("D1") = "" Then
        int_tags = 1
    Else
        int_tags = RangeEnd(Range("C1"), xlToRight).count
    End If
        
    int_rows = Range("A1")
    
    ''delete the formulas for whatever is in the sheet now
    Dim rng_array As Range
    Set rng_array = Intersect(RangeEnd(Range("B4"), xlDown, xlToRight), ActiveSheet.UsedRange)
    If Not rng_array Is Nothing Then
        rng_array.Formula = ""
    End If
    
    'LOGIC
    'for the first tag, create the formula for 2 columns
    'for the later tags, create formula in one column
    
    'formula for first column
    Range(Range("B4"), Range("B4").Offset(int_rows, 1)).FormulaArray = _
        "=PISampDat(R1C3,R2C1,R2C2,R3C2,1,""CPAMHCC-PIMS01"")"
        
    Dim int_col As Integer
    For int_col = 1 To int_tags - 1
    
        'for the first columnn, put the forumula down.  copy right for the others.
        If int_col = 1 Then
            Range(Range("C4").Offset(, int_col), Range("C4").Offset(int_rows, int_col)).FormulaArray = _
                "=PISampDat(D$1,$A$2,$B$2,$B$3,0,""CPAMHCC-PIMS01"")"
        Else
            Range(Range("C4").Offset(, int_col), Range("C4").Offset(int_rows, int_col)).FillRight
        End If
    
    Next int_col
    
    'format dates
    RangeEnd(Range("B4"), xlDown).NumberFormat = "mm/dd/yyyy HH:MM"

End Sub

Public Sub PiDataPull()

    Dim corner As Range
    Dim sht As Worksheet
    
    Set sht = ActiveSheet
    Set corner = sht.Range("A1")
    
    'Add in text
    sht.Range("A1").Formula = "=24*6"
    sht.Range("A2").Formula = "=B2-B1"
    sht.Range("B1") = 1
    sht.Range("B2") = Now
    sht.Range("A3") = "Interval"
    sht.Range("B3").Formula = "=(B2-A2)*24*60/A1&""m"""
    
    sht.Range("A2:B2").NumberFormat = "mm/dd/yyyy HH:MM"
    
    'get the pi tags
    Dim rng_tags As Range
    Set rng_tags = Range(corner, corner.End(xlToRight)).Offset(, 2).SpecialCells(xlCellTypeConstants)
    
    rng_tags.Offset(1).FormulaR1C1 = "=PITagAtt(R[-1]C,""descriptor"")"
    rng_tags.Offset(2).FormulaR1C1 = "=PITagAtt(R[-2]C,""engunits"")"

End Sub
