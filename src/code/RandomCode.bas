Attribute VB_Name = "RandomCode"
'---------------------------------------------------------------------------------------
' Module    : RandomCode
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains a lot of junk code that was stored.  Most is too specific to be useful.
'---------------------------------------------------------------------------------------




'''this one goes through a data source and alphabetizes it.
'''keeping mainly for the select case and find/findnext
Sub AlphabetizeAndReportWithDupes()

    Dim rng_data As Range
    Set rng_data = Range("B2:B28")

    Dim rng_output As Range
    Set rng_output = Range("I2")

    Dim arr As Variant
    arr = Application.Transpose(rng_data.Value)
    QuickSort arr
    'arr is now sorted

    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        
        'if duplicate, use FindNext, else just Find
        Dim rng_search As Range
        Select Case True
            Case i = LBound(arr), UCase(arr(i)) <> UCase(arr(i - 1))
                Set rng_search = rng_data.Find(arr(i))
            Case Else
                Set rng_search = rng_data.FindNext(rng_search)
        End Select

        ''''do your report stuff in here for each row
        'copy data over
        rng_output.Offset(i - 1).Resize(, 6).Value = rng_search.Resize(, 6).Value

    Next i
End Sub


Sub Rand_OpenFilesAndCopy()

    Dim sht_data As Worksheet
    Dim sht_output As Worksheet
    
    Set sht_output = ActiveSheet

    Dim path As Variant
    Dim folder As Variant
    
    Application.ScreenUpdating = False
    ' Another static folder
    folder = "O:\HCCShare\Operations\PE\Plant 8\Production Engineer\BWall\2013 11 Rheo troubleshooting\Recipes\PE7\2\"
    
    path = Dir(folder)
    
    Do While path <> ""

        Dim wkbk As Workbook
        Set wkbk = Workbooks.Open(folder & path)
        Set sht_data = wkbk.Sheets(1)
        sht_data.UsedRange.Copy
        
        sht_output.Cells(sht_output.UsedRange.Rows.count + 1, 1) = wkbk.name
        sht_output.Cells(sht_output.UsedRange.Rows.count, 2).PasteSpecial xlPasteValues
        
        wkbk.Close False
        
        path = Dir
    
    Loop

End Sub


Sub Rand_PrintMultiple()

    'go through the tags, pick one, put it in place
    
    'print out a PDF to a file
    
    Application.ScreenUpdating = False
    'Another static folder
    Dim rng_tag As Range
    Dim str_path As String
    str_path = "C:\Documents and Settings\wallbd\Application Data\PDF OUTPUT\"
    
    For Each rng_tag In Range("TAGS[TAG]").SpecialCells(xlCellTypeVisible)
        
        Range("C1") = rng_tag
        
        Sheets("SUMMARY").ExportAsFixedFormat xlTypePDF, str_path & rng_tag & ".PDF", , , , , , False
        
        'code is used to get a summary
        'Dim sht As Worksheet
        'Set sht = Sheets("ALL TAGS")
        
        'sht.Range("A1").EntireRow.Insert
        'sht.Range("A1") = rng_tag
        
        'Range("I8:L8").Copy
        'sht.Range("B1").PasteSpecial xlPasteValues
        'sht.Range("F1").Value = str_path & rng_tag & ".PDF"
    
    Next rng_tag
    
    Application.ScreenUpdating = True

End Sub

Sub Rand_PrintMultiplePvVsOp()

    'go through the tags, pick one, put it in place
    
    'print out a PDF to a file
    
    Application.ScreenUpdating = False
    'Another static folder
    Dim rng_tag As Range
    Dim str_path As String
    str_path = "C:\Documents and Settings\wallbd\Application Data\PDF OUTPUT\"
    
    For Each rng_tag In Range("tag_table[TAG]").SpecialCells(xlCellTypeVisible)
        
        Range("Charts!C3") = rng_tag
        
        Sheets("CHARTS").ExportAsFixedFormat xlTypePDF, str_path & rng_tag & "-" & rng_tag.Offset(, 5) & ".PDF", , , , , , False
    
    Next rng_tag
    
    Application.ScreenUpdating = True

End Sub

Function Download_File(ByVal vWebFile As String, ByVal vLocalFile As String) As Boolean
    Dim oXMLHTTP As Object, i As Long, vFF As Long, oResp() As Byte

    'You can also set a ref. to Microsoft XML, and Dim oXMLHTTP as MSXML2.XMLHTTP
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    oXMLHTTP.Open "GET", vWebFile, False 'Open socket to get the website
    oXMLHTTP.Send 'send request

    'Wait for request to finish
    Do While oXMLHTTP.readyState <> 4
    DoEvents
    Loop

    oResp = oXMLHTTP.responseBody 'Returns the results as a byte array

    'Create local file and save results to it
    vFF = FreeFile
    If Dir(vLocalFile) <> "" Then Kill vLocalFile
    Open vLocalFile For Binary As #vFF
    Put #vFF, , oResp
    Close #vFF

    'Clear memory
    Set oXMLHTTP = Nothing
End Function

Sub Rand_DownloadFromSheet()

    Dim rng_addr As Range
    
    Dim str_folder As Variant
    'Another static folder
    str_folder = "C:\Documents and Settings\wallbd\Application Data\DSP Guide\"
    
    For Each rng_addr In Range("B2:B35")
    
        Download_File rng_add, str_folder & rng_addr.Offset(, 1)
    
    Next rng_addr

End Sub

Sub Rand_CommonPrintSettings()

Application.ScreenUpdating = False
Dim sht As Worksheet

For Each sht In Sheets
    sht.PageSetup.PrintArea = ""
    sht.ResetAllPageBreaks
    sht.PageSetup.PrintArea = ""
    
    With sht.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.75)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Next sht
    
    Application.ScreenUpdating = True
End Sub


Sub Rand_DumpTextFromAllSheets()

    Dim c As Range
    Dim s As Worksheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim main As Workbook
    Set main = ActiveWorkbook
    
    Dim w As Workbook
    Dim sw As Worksheet
    
    Set w = Application.Workbooks.Add
    Set sw = w.Sheets.Add
    
    Dim Row As Integer
    Row = 0
    For Each s In main.Sheets
        For Each c In s.UsedRange.SpecialCells(xlCellTypeConstants)
            sw.Range("A1").Offset(Row) = c
            Row = Row + 1
        Next c
    Next s

End Sub


Sub Rand_ApplyHeadersAndFootersToAll()

    Dim sht As Worksheet
    Dim sht_hdr As Worksheet
    
    Set sht_hdr = ActiveSheet
    
    For Each sht In Sheets
        sht.PageSetup.LeftHeader = sht_hdr.PageSetup.LeftHeader
        sht.PageSetup.CenterHeader = sht_hdr.PageSetup.CenterHeader
        sht.PageSetup.RightHeader = sht_hdr.PageSetup.RightHeader
        sht.PageSetup.LeftFooter = sht_hdr.PageSetup.LeftFooter
        sht.PageSetup.CenterFooter = sht_hdr.PageSetup.CenterFooter
        sht.PageSetup.RightFooter = sht_hdr.PageSetup.RightFooter
    Next sht

End Sub

'Takes a table of values and flattens it.
Sub Rand_Matrix()

    Dim rng_left As Range
    Dim rng_top As Range
    Dim rng_body As Range
        
    Set rng_left = Application.InputBox("Select left column", Type:=8)
    Set rng_top = Application.InputBox("Select top column", Type:=8)
    
    Dim int_left As Integer, int_top As Integer
    
    Set rng_body = Range(Cells(rng_left.Row, rng_top.Column), _
                            Cells(rng_left.Rows(rng_left.Rows.count).Row, rng_top.Columns(rng_top.Columns.count).Column))
                            
    Dim sht_out As Worksheet
    Set sht_out = Application.Worksheets.Add()
    
    Dim rng_cell As Range
    
    Dim int_row As Integer
    int_row = 1
    
    For Each rng_cell In rng_body.SpecialCells(xlCellTypeConstants)
        sht_out.Range("A1").Offset(int_row) = rng_left.Cells(rng_cell.Row - rng_left.Row + 1, 1)
        sht_out.Range("B1").Offset(int_row) = rng_top.Cells(1, rng_cell.Column - rng_top.Column + 1)
        sht_out.Range("C1").Offset(int_row) = rng_cell
        
        int_row = int_row + 1
    Next rng_cell

End Sub

Sub Rand_CopyPasteValuesIntoNewSheet()

    Dim sht_new As Worksheet
    Dim sht_current As Worksheet
    
    Set sht_current = ActiveSheet
    
    Set sht_new = Worksheets.Add
    sht_current.UsedRange.Copy
    sht_new.PasteSpecial xlPasteValuesAndNumberFormats
    

End Sub

Sub Rand_ConvertToString()

    Dim cell As Range
    Dim sel As Range
    
    Set sel = Selection
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each cell In Intersect(sel, sel.Parent.UsedRange)
        If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            cell.Value = CStr(cell.Value)
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Sub Rand_KeepCellsWithText()

    Selection.SpecialCells(xlCellTypeConstants).Select

End Sub

Sub Rand_DeleteHiddenSheets()

    Dim sht As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each sht In Worksheets
        If sht.Visible = xlSheetHidden Then
            sht.Delete
        End If
    Next sht
    
    Application.DisplayAlerts = True

End Sub
