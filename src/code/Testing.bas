Attribute VB_Name = "Testing"
Option Explicit

Public Sub ComputeDistanceMatrix()

'get the range of inputs, along with input name
    Dim inputRange As Range
    Set inputRange = Application.InputBox("Select input data", "Input", Type:=8)

    'Dim myRange_ID As Range
    'Set myRange_ID = Application.InputBox("Select ID data", "ID", Type:=8)

    'turning off updates makes a huge difference here... could also use array for output
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    'create new workbook
    Dim myBook As Workbook
    Set myBook = Workbooks.Add

    Dim newSheet As Worksheet
    Set newSheet = myBook.Sheets(1)
    newSheet.name = "scaled data"

    'copy data over to standardize
    inputRange.Copy myBook.Sheets(1).Range("A1")

    'go to edge of data, add a column, add STANDARDIZE, copy paste values, delete
    
    Dim dataRange As Range
    Set dataRange = newSheet.Range("A1").CurrentRegion

    Dim myColumn As Range
    For Each myColumn In dataRange.Columns

        'edge cell
        Dim edgeCell As Range
        Set edgeCell = newSheet.Cells(1, newSheet.Columns.count).End(xlToLeft).Offset(, 1)
        
        'do a normal dist standardization
        '=STANDARDIZE(A1,AVERAGE(A:A),STDEV.S(A:A))
        
        edgeCell.Formula = "=IFERROR(STANDARDIZE(" & myColumn.Cells(1, 1).Address(False, False) & ",AVERAGE(" & _
            myColumn.Address & "),STDEV.S(" & myColumn.Address & ")),0)"
        
        'do a simple value over average to detect differences
        edgeCell.Formula = "=IFERROR(" & myColumn.Cells(1, 1).Address(False, False) & "/AVERAGE(" & _
            myColumn.Address & "),1)"
            
        'fill that down
        Range(edgeCell, edgeCell.Offset(, -1).End(xlDown).Offset(, 1)).FillDown

    Next
    
    Application.Calculate
    newSheet.UsedRange.Value = newSheet.UsedRange.Value
    dataRange.EntireColumn.Delete
    
    Dim distanceSheet As Worksheet
    Set distanceSheet = myBook.Worksheets.Add()
    distanceSheet.name = "distances"

    Dim outboundRange As Range
    Set outboundRange = distanceSheet.Range("A1")

    'loop through each row with each other row
    Dim firstRow As Range
    Dim secondRow As Range
    
    Set inputRange = newSheet.Range("A1").CurrentRegion

    For Each firstRow In inputRange.Rows
        For Each secondRow In inputRange.Rows

            'loop through each column and compute the distance
            Dim squaredDistance As Double
            squaredDistance = 0

            Dim currentColumn As Long
            For currentColumn = 1 To firstRow.Cells.count
                squaredDistance = squaredDistance + (firstRow.Cells(1, currentColumn) - secondRow.Cells(1, currentColumn)) ^ 2
            Next

            'take the sqrt of that value and output
            outboundRange.Value = squaredDistance ^ 0.5

            'get to next column for output
            Set outboundRange = outboundRange.Offset(, 1)
        Next

        'drop down a row and go back to left edge
        Set outboundRange = outboundRange.Offset(1).End(xlToLeft)
    Next

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    distanceSheet.UsedRange.NumberFormat = "0.00"
    distanceSheet.UsedRange.EntireColumn.AutoFit
    
    'do the coloring
    Formatting_AddCondFormat distanceSheet.UsedRange

End Sub

Sub RemoveAllLegends()

    Dim myChartObject As ChartObject
    
    For Each myChartObject In Chart_GetObjectsFromObject(Selection)
        myChartObject.Chart.HasLegend = False
        myChartObject.Chart.HasTitle = True
        
        myChartObject.Chart.SeriesCollection(1).MarkerSize = 4
    Next

End Sub

Sub ApplyFormattingToEachColumn()
    Dim myRange As Range
    For Each myRange In Selection.Columns

        Formatting_AddCondFormat myRange
    Next
End Sub

Private Sub Formatting_AddCondFormat(ByVal myRange As Range)

        myRange.FormatConditions.AddColorScale ColorScaleType:=3
        myRange.FormatConditions(myRange.FormatConditions.count).SetFirstPriority
        myRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
        With myRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .Color = 7039480
            .TintAndShade = 0
        End With
        myRange.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
        myRange.FormatConditions(1).ColorScaleCriteria(2).Value = 50
        With myRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .Color = 8711167
            .TintAndShade = 0
        End With
        myRange.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .Color = 8109667
            .TintAndShade = 0
        End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : TraceDependentsForAll
' Author    : @byronwall
' Date      : 2015 11 09
' Purpose   : Quick Sub to iterate through Selection and Trace Dependents for all
'---------------------------------------------------------------------------------------
'
Sub TraceDependentsForAll()

    Dim myRange As Range
    
    For Each myRange In Intersect(Selection, Selection.Parent.UsedRange)
        myRange.ShowDependents
    Next myRange

End Sub

