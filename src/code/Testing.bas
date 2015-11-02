Attribute VB_Name = "Testing"
Option Explicit

Sub SeriesSplitIntoBins()

    On Error GoTo ErrorNoSelection

    Dim rngSelection As Range
    Set rngSelection = Application.InputBox("Select category range with heading", _
                                            Type:=8)
    Set rngSelection = Intersect(rngSelection, _
                                 rngSelection.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + _
                                                                                                xlNumbers + xlTextValues)

    Dim rngValues As Range
    Set rngValues = Application.InputBox("Select values range with heading", _
                                         Type:=8)
    Set rngValues = Intersect(rngValues, rngValues.Parent.UsedRange)

    ''need to prompt for max/min/bins
    Dim dbl_max As Double, dbl_min As Double, int_bins As Integer

    dbl_min = Application.InputBox("Minimum value.", "Min", _
                                   WorksheetFunction.Min(rngSelection), Type:=1)
    dbl_max = Application.InputBox("Maximum value.", "Max", _
                                   WorksheetFunction.Max(rngSelection), Type:=1)
    int_bins = Application.InputBox("Number of groups.", "Bins", _
                                    WorksheetFunction.RoundDown(Math.Sqr(WorksheetFunction.count(rngSelection)), 0), _
                                    Type:=1)

    On Error GoTo 0

    'determine default value
    Dim strDefault As Variant
    strDefault = Application.InputBox("Enter the default value", "Default", "#N/A")

    'detect cancel and exit
    If StrPtr(strDefault) = 0 Then
        Exit Sub
    End If

    ''TODO prompt for output location

    rngValues.EntireColumn.Offset(, 1).Resize(, int_bins + 2).Insert
    'head the columns with the values

    ''TODO add a For loop to go through the bins

    Dim int_binNo As Integer
    For int_binNo = 0 To int_bins
        rngValues.Cells(1).Offset(, int_binNo + 1) = dbl_min + (dbl_max - dbl_min) * int_binNo / int_bins
    Next

    'add the last item
    rngValues.Cells(1).Offset(, int_bins + 2).FormulaR1C1 = "=RC[-1]"

    ''TODO add formulas for first, mid, last columns
    'FIRST =IF($D2 <=V$1,$U2,#N/A)
    '=IF(RC4 <=R1C,RC21,#N/A)

    'MID =IF(AND($D2 <=W$1, $D2>V$1),$U2,#N/A)  '''W current, then left
    '=IF(AND(RC4 <=R1C, RC4>R1C[-1]),RC21,#N/A)

    'LAST =IF($D2>AA$1,$U2,#N/A)
    '=IF(RC4>R1C[-1],RC21,#N/A)

    ''TODO add number format to display header correctly (helps with charts)

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim strFormula As Variant
    strFormula = "=IF(AND(RC" & rngSelection.Column & _
               " <=R" & rngValues.Cells(1).Row & "C," & _
                 "RC" & rngSelection.Column & ">R" & rngValues.Cells(1).Row & "C[-1]" & _
                 ")" & _
                 ",RC" & rngValues.Column & "," & strDefault & ")"

    Dim str_FirstFormula As Variant
    str_FirstFormula = "=IF(AND(RC" & rngSelection.Column & _
                     " <=R" & rngValues.Cells(1).Row & "C)" & _
                       ",RC" & rngValues.Column & "," & strDefault & ")"

    Dim str_LastFormula As Variant
    str_LastFormula = "=IF(AND(RC" & rngSelection.Column & _
                    " >R" & rngValues.Cells(1).Row & "C)" & _
                      ",RC" & rngValues.Column & "," & strDefault & ")"

    Dim rngFormula As Range
    Set rngFormula = rngValues.Offset(1, 1).Resize(rngValues.Rows.count - 1, _
                                                   int_bins + 2)
    rngFormula.FormulaR1C1 = strFormula

    'override with first/last
    rngFormula.Columns(1).FormulaR1C1 = str_FirstFormula
    rngFormula.Columns(rngFormula.Columns.count).FormulaR1C1 = str_LastFormula

    rngFormula.EntireColumn.AutoFit
    
    'set the number formats
    rngFormula.Offset(-1).Rows(1).Resize(1, int_bins + 1).NumberFormat = "<= General"
    rngFormula.Offset(-1).Rows(1).Offset(, int_bins + 1).NumberFormat = "> General"

    Exit Sub

ErrorNoSelection:
    'TODO: consider removing this prompt
    MsgBox "No selection made.  Exiting.", , "No selection"

End Sub

