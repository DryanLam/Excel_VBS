

VBA

Queue: https://www.youtube.com/user/ExcelCampus/playlists


Power Pivot Table: 1 hour https://www.youtube.com/watch?v=ohGFPF12Qwcs
icon: http://xahlee.info/comp/unicode_punctuation_symbols.html

!!! Sheet >> Insert Module to use Function as formular

power query: Copy multiple file to 1
https://www.youtube.com/watch?v=9glqnywsL60

Alt + F11: Open VBA
F8: Next line
F5: Run VBA


Save As: Excel Macro-Enabled Workbook

Exclude Macro: 
Save As >> Workwork -> Yes

Export & Import
Module >> Right click >> Export
File >> Import

Insert Module
VBA project >> Sheet1 >> Insert >> Module


BUTTON
Developer >> Insert >> Button
Insert any Image >> Right click >> Assign Macro

= _ : Giúp liên kết các dòng code trong VBA

--- Hierarchy
Application
    |
Workbook
    |
Worksheet
    |
Range
    |
Cells

------------------------- COMMON EXCEL FUNCTION
=INDEX(array, row_num, [column_num])

    A   B
1   ID 	name
2   1	A
3   2	B

=INDEX(A2:B2&A3:B3,0,1) >> {12,AB}

>> 12

=MATCH(lookup_value,lookup_array,[match_type])
0: exact
return index_number in array



---
=char(50)   : icon up

=char(60)   : icon down

=char(44)   : icon score
---
Gantt: 
Format condition:
	+ Custom
	+ Use a formular
	
>> if(AND(E$2 >= $C3, E$2<= $C3 + $D3 -1), true, flase)
	E2: Firt day in gantt
	C3: Start date
	D3: Duration

------------------------- COMMON VBA - https://www.youtube.com/watch?v=DT0QOoLvM10
workbooks("Name").Sheets("Name")...

Range(Cells(1,2),Cells(2,7)).Clear
Range(Cells(1,2),Cells(2,7)).Delete
Range(Cells(1,2),Cells(2,7)) = "hello"

---
Sub thangSelect()
    Sheets("Thang 1").Select
End Sub
---
Call Excel Function in VBA:
Excel.WorksheetFunction.<FunctionName>(fvalue)
    >> v = Excel.WorksheetFunction.Sum(Range(A2:B6))




------------------------- Turn off alert
Application.DisplayAlerts = False
------------------------- Delete Sheet
ThisWorkbook.Sheet("Name").Delete

Sheets(1).Delete

Application.DisplayAlerts = False
For Each ws In ThisWorkbook.Sheets
 If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
    ws.Delete
 End if
Next



------------------------- Print
ThisWorkbook.Sheet("Name").PrintOut preview = False
ThisWorkbook.Sheet("Name").PrintOut preview = False

------------------------- OPEN & CLOSE Workbooks
Sub OpenWorkbook(name As String)
    Workbooks.Open name
End Sub

Sub CloseWorkbook(name As String)
    Workbooks(name).Close SaveChanges:=True
End Sub

------------------------- SUM All Range
Sub tongRange()
    Dim tong As Long
    tong = 0
    For Each i in Cells(1,1).CurrentRegion
        tong = tong + i
    Next
    MsgBox tong
End Sub

------------------------- offset - index - match
= offset(index($A$3:$A$12,MATCH($H$2,$A$3:$A$12,0)),0,1,COUNTIF($$A$3:$A$12,$H$2)4)

H2: value filter
$A$3:$A$12: range of filter value
4: 4 colums getting

---
Workbooks("New Data.xlsx").Worksheets("Export").Range("A2:D9").Copy
---
Cells(rowIndex, colIndex).Value = "Hi"

-------------------------  Last row, last column
Dim lngRow As Long
lngRow = Cells(Rows.Count, "A").End(xlUp).Row

LastRowColA = Range("F" & Rows.Count).End(xlUp).Row

lColumn = Cells(1, Columns.Count).End(xlToLeft).Column

'Last column number including blank
'lcSouce = srcSheet.Range("XFD1").End(xlToLeft).Column
'Column name
lcDestNum = srcSheet.Range("A1").CurrentRegion.Columns.Count
lcSource = Split(Cells(1, lcDestNum).Address, "$")(1)


---
Sub test_03()
    Range("A2, C2:F2,A5:B7") = 999
End Sub

' Use formular
Sub test_04()
    'Use formular
    Range("C2:F4") = _
    "=RANDBETWEEN(10,100)"
    
    'This is to hide formular
    Range("C2:F4") = Range("C2:F4").Value
    
    'Range SUM -> Just need firt column
    Range("B6") = "Total"
    Range("C6:F6") = "=SUM(C2:C4)"
    
    'Total SUM
    Range("B9") = "Grand Total"
    Range("C9") = "=SUM(C2:F4)"
End Sub


Sub check_file_exist()
    Dim file_path As String
    file_path = "C:\"
    
    If Dir(file_path & "dunglam.xlsx") = "" Then
        MsgBox "File not found!!!"
    End If
End Sub

Sub tao_file_excel()
    Dim wb As Workbook
    
    Dim file_path As String
    file_path = "D:\"
    
    Set wb = Workbooks.Add
    
    wb.Activate
    wb.SaveAs file_path & "taofile.xlsx"
End Sub



Sub filter_region()
    Dim sh As Worksheet
    Dim rng As Range
    Dim lr  As Long
    
    
    'Point to the working sheet
    'Set sh = thisworkbook.Worksheets("SheetName")
    Set sh = Sheet1
    
    'Last row
    lr = sh.Range("H" & Rows.Count).End(xlUp).Row
    
    
    'Range active value
    Set rng = sh.Range("H4:J" & lr)
    
    'Filter Region: 2nd column
    'Criteria1: filter cell of data validation list
    rng.AutoFilter field:=2, Criteria1:=sh.Range("I2").Value
End Sub



----------------------------------------------------------------------- Vlookup & Join value with ","
Sheet1 >> Insert Module

Function mVlookUp(lookup_value, lookup_range As Range, index_col As Long)
    result = ""
    For Each r In lookup_range
        If r = lookup_value Then
            result = result & "," & r.Offset(0, index_col - 1)
        End If
    Next
    
    mVlookUp = Right(result, Len(result) - 1)
End Function

----------------------------------------------------------------------- Filter Advance
=COUNTIF(tblFiltetList[tableName],[@Colname])=1

---
Sub filterpage()
    nfilter = Array("AA", "BB", "CC")
    
    For i = LBound(nfilter) To UBound(nfilter)
        ' Create sheet and naming
        Set sh = Sheets.Add
        sh.Name = nfilter(i)
        
        'Init Filter: col B2. Build temp cell W2 to store filter value
        Sheet1.Range("W2").Formula = "=B2=""" & nfilter(i) & """"
        
        Sheet1.Range("A:C").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Sheet1.Range("W1:W2"), CopyToRange:=sh.Range("A1:C1"), _
        Unique:=False
    Next i
End Sub

---

Sub filterValues()
        ' Create sheet and naming
        Set sh = Sheets.Add
        sh.Name = "New Data"
        
        'Init Filter: col B2. Build temp cell W2 to store filter value
        Sheet1.Range("W2").Formula = "=AND(A2>1,A2<9)"
        
        Sheet1.Range("A:C").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Sheet1.Range("W1:W2"), CopyToRange:=sh.Range("A1:C1"), _
        Unique:=False
End Sub

---
Sub filter_by_multi()
    Dim cn As Range, tc As Range
    
    For Each cn In Sheet1.Range("I2:I5")
        For Each tc In Sheet1.Range("J2:J3")
            With Sheet1.Range("E:G")
                ' Turn off autofilter
                .Parent.AutoFilterMode = False
                
                ' Turn new filter
                .AutoFilter
                
                ' Filter field 1(Chi nhanh) with criteria1 is cn
                .AutoFilter Field:=1, Criteria1:=cn
                ' Filter field 3(Tieu chi) with criteria1 is tc
                .AutoFilter Field:=3, Criteria1:=tc
                ' Copy data filter
                .Parent.AutoFilter.Range.Copy
                
                'Create sheet after current sheet1
                Sheets.Add after:=Sheets(Sheets.Count)
                'Naming
                Sheets(Sheets.Count).Name = cn & "-" & tc
                'Paste filter data from A1 of new sheet
                Sheets(Sheets.Count).Range("A1").PasteSpecial xlPasteValues
            End With
        Next
    Next
End Sub

---
Filter data from multi conditions
' This can filter letter contains, date, whatelse
'Filter criteria
'    A       B
' 1  Ho      Nam sinh
' 2  L       > 2010
Sub filterMulti()
    Dim data_rg, filters_rg,copy_rg As Range

    'Init data
    Set data_rg = Sheets("Data").range("A1").CurrentRegion
    'Init filter criteria
    Set filters_rg =  Sheets("Criteria").range("A1").CurrentRegion
    'Init dest
    Set copy_rg = Sheets("Copy").Range("A1")

    'Clear data before return result
    Sheets("Data").range("A:Z").Delete

    'Resulting filter data to sheet
    data_rg.AdvancedFilter xlFilterCopy, filters_rg, copy_rg
End Sub

----------------------------------------------------------------------- TO XML
Sub MakeXML(iCaptionRow As Integer, 
            iDataStartRow As Integer, 
            sOutputFileName As String)
    Dim Q As String
    'Double quotes
    Q = Chr$(34)

    Dim sXML As String

    sXML = "<?xml version=" & Q & "1.0" & Q & " encoding=" & Q & "UTF-8" & Q & "?>"
    sXML = sXML & "<rows>"


    ''--determine count of columns
    Dim iColCount As Integer
    iColCount = 1
    While Trim$(Cells(iCaptionRow, iColCount)) > ""
        iColCount = iColCount + 1
    Wend

    Dim iRow As Integer
    iRow = iDataStartRow

    While Cells(iRow, 1) > ""
        sXML = sXML & "<row id=" & Q & iRow & Q & ">"

        For icol = 1 To iColCount - 1
           sXML = sXML & "<" & Trim$(Cells(iCaptionRow, icol)) & ">"
           sXML = sXML & Trim$(Cells(iRow, icol))
           sXML = sXML & "</" & Trim$(Cells(iCaptionRow, icol)) & ">"
        Next

        sXML = sXML & "</row>"
        iRow = iRow + 1
    Wend
    sXML = sXML & "</rows>"

    Dim nDestFile As Integer, sText As String

    ''Close any open text files
    Close

    ''Get the number of the next free text file
    nDestFile = FreeFile

    ''Write the entire file to sText
    Open sOutputFileName For Output As #nDestFile
    Print #nDestFile, sXML
    Close
End Sub

Sub test()
    MakeXML 1, 2, "C:\Users\jlynds\output2.xml"
End Sub
--- cooking 
Sub MakeXML(iCaptionRow As Integer, iDataStartRow As Integer, sOutputFileName As String)
    Dim Q As String
    'Double quotes
    Q = Chr$(34)
    'Line Feed
    Lf = Chr$(10)
    'Carriage Line
    Cl = Chr$(13)
    'Equal
    Eq = Chr$(61)
    'Space
    Sp = Chr$(32)
    'Null
    Nl = Chr$(0)

    Dim sXML As String

    sXML = "<?xml version=" & Q & "1.0" & Q & " encoding=" & Q & "UTF-8" & Q & "?>"
    sXML = sXML & Lf & "<rows>"


    ''--determine count of columns
    Dim iColCount As Integer
    iColCount = 1
    While Trim$(Cells(iCaptionRow, iColCount)) > ""
        iColCount = iColCount + 1
    Wend

    Dim iRow As Integer
    iRow = iDataStartRow

    While Cells(iRow, 1) > ""
        sXML = sXML & Lf & vbTab & "<row id=" & Q & iRow & Q

        For icol = 1 To iColCount - 1
           sXML = sXML & Lf & vbTab & vbTab & Trim$(Cells(iCaptionRow, icol)) & Eq
           sXML = sXML & Trim$(Cells(iRow, icol)) & Sp
        Next

        sXML = sXML & "/>"
        iRow = iRow + 1
    Wend
    sXML = sXML & Lf & "</rows>"

    Dim nDestFile As Integer, sText As String

    ''Close any open text files
    Close

    ''Get the number of the next free text file
    nDestFile = FreeFile

    ''Write the entire file to sText
    Open sOutputFileName For Output As #nDestFile
    Print #nDestFile, sXML
    Close
End Sub

Sub test()
    MakeXML 1, 2, "D:\excel_parser.xml"
End Sub

---
Sub MakeXML()
    
    sOutputFileName = "D:\excel_parser.xml"

    Q = Chr$(34)
    Lf = Chr$(10)
    Cl = Chr$(13)
    Eq = Chr$(61)
    Sp = Chr$(32)
    Oq = Chr$(60)
    Cq = Chr$(62)
    Sl = Chr$(47)
    
    Dim sXML As String
    'XML Template
    sXML = "<?xml version=" & Q & "1.0" & Q & " encoding=" & Q & "UTF-8" & Q & "?>"
    'OPEN DATAPACKET
    sXML = sXML & Lf & Oq & Trim$(Cells(1, 1)) & Sp & Trim$(Cells(1, 2)) & Eq & Q & Trim$(Cells(2, 2)) & Q & Cq
    'OPEN METADATA
    sXML = sXML & Lf & Sp & Sp & Oq & Trim$(Cells(1, 3)) & Cq
    'OPEN FIELDS
    sXML = sXML & Lf & vbTab & Oq & Trim$(Cells(1, 4)) & Cq
    iRow = 2
    'FIELD
    While Cells(iRow, 6) > ""
        sXML = sXML & Lf & vbTab & Sp & Sp & Oq & Trim$(Cells(1, 5))
        For iCol = 6 To 9
            If Trim$(Cells(iRow, iCol)) > "" Then
                sXML = sXML & Lf & vbTab & Sp & Sp & Sp & Sp & Trim$(Cells(1, iCol)) & Eq & Q
                sXML = sXML & Trim$(Cells(iRow, iCol)) & Q
            End If
        Next
        sXML = sXML & Sp & Sl & Cq
        iRow = iRow + 1
    Wend
    'CLOSE FIELDS
    sXML = sXML & Lf & vbTab & Oq & Sl & Trim$(Cells(1, 4)) & Cq
    'PARAMS
    sXML = sXML & Lf & vbTab & Oq & Trim$(Cells(1, 10))
    sXML = sXML & Lf & vbTab & Sp & Sp & Trim$(Cells(1, 11)) & Eq & Q & Trim$(Cells(2, 11)) & Q
    sXML = sXML & Lf & vbTab & Sp & Sp & Trim$(Cells(1, 12)) & Eq & Q & Trim$(Cells(2, 12)) & Q
    sXML = sXML & Sp & Sl & Cq
    'CLOSE METADATA
    sXML = sXML & Lf & Sp & Sp & Oq & Sl & Trim$(Cells(1, 3)) & Cq
    'OPEN ROWDATA
    sXML = sXML & Lf & Sp & Sp & Oq & Trim$(Cells(1, 13)) & Cq
    'ROW
    Dim mRow As Integer
    mRow = 2
    While Cells(mRow, 15) > ""
        sXML = sXML & Lf & vbTab & Oq & Trim$(Cells(1, 14))
        For mCol = 15 To 21
            sXML = sXML & Lf & vbTab & Sp & Sp & Sp & Sp & Trim$(Cells(1, mCol)) & Eq & Q
            sXML = sXML & Trim$(Cells(mRow, mCol)) & Q
        Next
        sXML = sXML & Sp & Sl & Cq
        mRow = mRow + 1
    Wend
    'CLOSE ROWDATA
    sXML = sXML & Lf & Sp & Sp & Oq & Sl & Trim$(Cells(1, 13)) & Cq
    'CLOSE DATAPACKET
    sXML = sXML & Lf & Oq & Sl & Trim$(Cells(1, 1)) & Cq
    
    
    Dim nDestFile As Integer, sText As String
    ''Close any open text files
    Close
    ''Get the number of the next free text file
    nDestFile = FreeFile
    ''Write the entire file to sText
    Open sOutputFileName For Output As #nDestFile
    Print #nDestFile, sXML
    Close
End Sub


---------------------------------------------------------- COPY  DATA FROM MULTTI-FILES TO OPEN
Sub copyFile()
Path = "D:\foldername\"
Filename = Dir(Path & "*.xls*")
Do While Filename <> ""
    Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
    For Each Sheet In ActiveWorkBook.Sheets
        Sheet.Copy after:=ThisWorkbook.Sheets(1)
    Next
Workbooks(Filename).Close
Filename = Dir()
Loop
End Sub

---------------------------------------------------------- COPY  DATA TO ANOTHER SHEET
Sub OpenWorkbook(name As String)
    Workbooks.Open name
End Sub

Sub CloseWorkbook(name As String)
    Workbooks(name).Close SaveChanges:=True
End Sub

Sub Copy_01()
    Worksheets("Sheet1").Range("A1:D9").Copy Worksheets("Sheet2").Range("A1")
End Sub

Sub CopyToNewSheet()
    ' Create sheet and naming
    Set nSheet = Sheets.Add
    nSheet.name = "Copy of Data " & Format(Now, "HHMMSS")
    
    Worksheets("Sheet1").Range("A1:D9").Copy nSheet.Range("A1")
End Sub

Sub CopyToExistingSheet()
    ' Init sheets
    Dim wb As Workbook
    Dim srcSheet, dstSheet As Worksheet
    Dim lrSource, lrDest As Long
    Dim lcSource As String
    
    Set wb = ThisWorkbook
    Set srcSheet = wb.Sheets("Sheet1")
    Set dstSheet = wb.Sheets("Copy of Data 111708")
    ' Source Sheet
    '' Last row
    lrSource = srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row
    '' Last col
    'Last column number including blank
    'lcSouce = srcSheet.Range("XFD1").End(xlToLeft).Column
    'Column name
    lcDestNum = srcSheet.Range("A1").CurrentRegion.Columns.Count
    lcSource = Split(Cells(1, lcDestNum).Address, "$")(1)
    
    ' Dest Sheet offset 1
    lrDest = dstSheet.Cells(dstSheet.Rows.Count, "A").End(xlUp).Offset(1).Row
    
    'Copy directly to dest
    'srcSheet.Range("A2:" & lcSource & lrSource).Copy dstSheet.Range("A" & lrDest)
    
    'Copy range to clipboard
    'srcSheet.Range("A2:" & lcSource & lrSource).Copy
    'PasteSpecial to paste values, formulas, formats, etc.
    'dstSheet.Range("A" & lrDest).PasteSpecial Paste:=xlPasteValues
End Sub

---------------------------------------------------------- SQL - ADODB connection
Function DATAFROMSQL(sqlStr)
    Dim t As Single
    t = Timer
    On Error GoTo errHandling
    Dim cnn As Object: Dim rst As Object
    Set cnn = CreateObject("ADODB.Connection"): Set rst = CreateObject("ADODB.Recordset")
    Const adUseClient = 3: Const adOpenStatic = 3: Const adLockReadOnly = 1
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    rst.Open sqlStr.Value, cnn, adOpenStatic, adLockReadOnly: rst.MoveFirst
    DATAFROMSQL = TransposeArray(rst.getrows())
    Sheet6.Range("A27") = "Success. Return " & rst.RecordCount & " records. " & (Timer - t) & " s"
    Exit Function
errHandling:
    Sheet6.Range("A27") = Err.Number & " | " & Err.Description
End Function
Private Function TransposeArray(myarray As Variant) As Variant
    Dim x As Long, y As Long, Xupper As Long, Yupper As Long, tempArray As Variant
    Xupper = UBound(myarray, 2)
    Yupper = UBound(myarray, 1)
    ReDim tempArray(Xupper, Yupper)
    For x = 0 To Xupper
        For y = 0 To Yupper
            tempArray(x, y) = myarray(y, x)
        Next y
    Next x
    TransposeArray = tempArray
End Function
Sub write_result()
On Error GoTo errHandling
    Dim rs: rs = DATAFROMSQL(Sheet6.Range("A2"))
    Sheet6.Range("L2:Z" & Sheet6.Rows.Count).ClearContents
    Sheet6.Range("L2").Resize(UBound(rs, 1) + 1, UBound(rs, 2) + 1).Value = rs
    Exit Sub
errHandling:
    Sheet6.Range("A27") = "No result or error in query"
End Sub

---------------------------------------------------------- Just temp
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adPersistXML = 1

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

''It wuld probably be better to use the proper name, but this is
''convenient for notes
strFile = Workbooks(1).FullName

''Note HDR=Yes, so you can use the names in the first row of the set
''to refer to columns, note also that you will need a different connection
''string for >=2007
strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFile _
        & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"


cn.Open strCon
rs.Open "Select * from [Sheet1$]", cn, adOpenStatic, adLockOptimistic

If Not rs.EOF Then
    rs.MoveFirst
    rs.Save "C:\Docs\Table1.xml", adPersistXML
End If

rs.Close
cn.Close

--- cooking 
Sub MakeXML(iCaptionRow As Integer, iDataStartRow As Integer, sOutputFileName As String)
    Dim Q As String
    'Double quotes
    Q = Chr$(34)
    'Line Feed
    Lf = Chr$(10)
    'Carriage Line
    Cl = Chr$(13)
    'Equal
    Eq = Chr$(61)
    'Space
    Sp = Chr$(32)
    'Null
    Nl = Chr$(0)

    Dim sXML As String

    sXML = "<?xml version=" & Q & "1.0" & Q & " encoding=" & Q & "UTF-8" & Q & "?>"
    sXML = sXML & Lf & "<rows>"


    ''--determine count of columns
    Dim iColCount As Integer
    iColCount = 1
    While Trim$(Cells(iCaptionRow, iColCount)) > ""
        iColCount = iColCount + 1
    Wend

    Dim iRow As Integer
    iRow = iDataStartRow

    While Cells(iRow, 1) > ""
        sXML = sXML & Lf & vbTab & "<row id=" & Q & iRow & Q

        For icol = 1 To iColCount - 1
           sXML = sXML & Lf & vbTab & vbTab & Trim$(Cells(iCaptionRow, icol)) & Eq
           sXML = sXML & Trim$(Cells(iRow, icol)) & Sp
        Next

        sXML = sXML & "/>"
        iRow = iRow + 1
    Wend
    sXML = sXML & Lf & "</rows>"

    Dim nDestFile As Integer, sText As String

    ''Close any open text files
    Close

    ''Get the number of the next free text file
    nDestFile = FreeFile

    ''Write the entire file to sText
    Open sOutputFileName For Output As #nDestFile
    Print #nDestFile, sXML
    Close
End Sub

Sub test()
    MakeXML 1, 2, "D:\excel_parser.xml"
End Sub



------------------------------------------- Helped
Sub MakeXML(iCaptionRow As Integer, iDataStartRow As Integer, sOutputFileName As String)
    Dim Q As String
    'Double quotes
    Q = Chr$(34)
    'Line Feed
    Lf = Chr$(10)
    'Carriage Line
    Cl = Chr$(13)
    'Equal
    Eq = Chr$(61)
    'Space
    sp = Chr$(32)
    'Null
    Nl = Chr$(0)
    '<
    Oq = Chr$(60)
    '>
    Cq = Chr$(62)
    '/
    Sl = Chr$(47)
    
    
    Dim sXML As String
    Dim iRow As Integer
    Dim iCol As Integer
    iRow = 2
    
    'Init
    sXML = "<?xml version=" & Q & "1.0" & Q & " encoding=" & Q & "UTF-8" & Q & "?>"
    
    'OPEN DATAPACKET
    sXML = sXML & Lf & Oq & Trim$(Cells(1, 1)) & sp & Trim$(Cells(1, 2)) & Eq & Q & Trim$(Cells(2, 2)) & Q & Cq

    'OPEN METADATA
    sXML = sXML & Lf & sp & sp & Oq & Trim$(Cells(1, 3)) & Cq
    
    'OPEN FIELDS
    sXML = sXML & Lf & vbTab & Oq & Trim$(Cells(1, 4)) & Cq
    ''--determine count of attr columns
    Dim attRows As Integer
    iColCount = Range("F" & Rows.Count).End(xlUp).Row + 1
    
    While Cells(iRow, 6) > ""
        sXML = sXML & Lf & vbTab & sp & sp & Oq & Trim$(Cells(1, 5))
        
        For iCol = 6 To 9
            If Trim$(Cells(iRow, iCol)) > "" Then
                sXML = sXML & Lf & vbTab & sp & sp & sp & sp & Trim$(Cells(1, iCol)) & Eq & Q
                sXML = sXML & Trim$(Cells(iRow, iCol)) & Q
            End If
        Next
        
        sXML = sXML & Sl & Cq
        iRow = iRow + 1
    Wend
    
    
    
    'CLOSE FIELDS
    sXML = sXML & Lf & vbTab & Oq & Sl & Trim$(Cells(1, 4)) & Cq
    'PARAMS
    sXML = sXML & Lf & vbTab & Oq & Trim$(Cells(1, 10))
    sXML = sXML & Lf & vbTab & sp & sp & Trim$(Cells(1, 11)) & Eq & Q & Trim$(Cells(2, 11)) & Q
    sXML = sXML & Lf & vbTab & sp & sp & Trim$(Cells(1, 12)) & Eq & Q & Trim$(Cells(2, 12)) & Q
    sXML = sXML & sp & Sl & Cq
    'CLOSE METADATA
    sXML = sXML & Lf & sp & sp & Oq & Sl & Trim$(Cells(1, 3)) & Cq
    
    'OPEN ROWDATA
    sXML = sXML & Lf & sp & sp & Oq & Trim$(Cells(1, 13)) & Cq
    'ROW
    Dim mRow As Integer
    mRow = 2
    While Cells(mRow, 15) > ""
        sXML = sXML & Lf & vbTab & Oq & Trim$(Cells(1, 14))
        
        For mCol = 15 To 21
            sXML = sXML & Lf & vbTab & sp & sp & sp & sp & Trim$(Cells(1, mCol)) & Eq & Q
            sXML = sXML & Trim$(Cells(mRow, mCol)) & Q
        Next
        
        sXML = sXML & Sl & Cq
        mRow = mRow + 1
    Wend
    'CLOSE ROWDATA
    sXML = sXML & Lf & sp & sp & Oq & Sl & Trim$(Cells(1, 13)) & Cq
    'CLOSE DATAPACKET
    sXML = sXML & Lf & Oq & Sl & Trim$(Cells(1, 1)) & Cq
    
    
    Dim nDestFile As Integer, sText As String
    ''Close any open text files
    Close

    ''Get the number of the next free text file
    nDestFile = FreeFile

    ''Write the entire file to sText
    Open sOutputFileName For Output As #nDestFile
    Print #nDestFile, sXML
    Close
End Sub

Sub test()
    MakeXML 1, 2, "D:\excel_parser.xml"
End Sub
---
Sub MakeXML()
    
    sOutputFileName = "D:\excel_parser.xml"

    Q = Chr$(34)
    Lf = Chr$(10)
    Cl = Chr$(13)
    Eq = Chr$(61)
    Sp = Chr$(32)
    Oq = Chr$(60)
    Cq = Chr$(62)
    Sl = Chr$(47)
    
    Dim sXML As String
    'XML Template
    sXML = "<?xml version=" & Q & "1.0" & Q & " encoding=" & Q & "UTF-8" & Q & "?>"
    'OPEN DATAPACKET
    sXML = sXML & Lf & "<DATAPACKET Version=""2.0"">"
    'OPEN METADATA
    sXML = sXML & Lf & "  <METADATA>"
    'OPEN FIELDS
    sXML = sXML & Lf & vbTab & "<FIELDS>"
    'FIELD
    sXML = sXML & Lf & vbTab & Sp & Sp & "<FIELD"
    sXML = sXML & Lf & vbTab & Sp & Sp & "  attrname=""STUFFCODE"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  fieldtype=""i4"" />"
    
    sXML = sXML & Lf & vbTab & Sp & Sp & "<FIELD"
    sXML = sXML & Lf & vbTab & Sp & Sp & "  attrname=""STUFFNAME"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  fieldtype=""string"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  WIDTH=""100"" />"
    
    sXML = sXML & Lf & vbTab & Sp & Sp & "<FIELD"
    sXML = sXML & Lf & vbTab & Sp & Sp & "  attrname=""BRCHNAME"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  fieldtype=""string"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  required=""true"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  WIDTH=""100"" />"
    
    sXML = sXML & Lf & vbTab & Sp & Sp & "<FIELD"
    sXML = sXML & Lf & vbTab & Sp & Sp & "  attrname=""CARDTIME"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  fieldtype=""SQLdateTime"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  required=""true"" />"
    
    sXML = sXML & Lf & vbTab & Sp & Sp & "<FIELD"
    sXML = sXML & Lf & vbTab & Sp & Sp & "  attrname=""CARDTYPEID"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  fieldtype=""i4"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  required=""true"" />"
    
    sXML = sXML & Lf & vbTab & Sp & Sp & "<FIELD"
    sXML = sXML & Lf & vbTab & Sp & Sp & "  attrname=""CLOCKID"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  fieldtype=""i4"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  required=""true"" />"
    
    sXML = sXML & Lf & vbTab & Sp & Sp & "<FIELD"
    sXML = sXML & Lf & vbTab & Sp & Sp & "  attrname=""IMGID"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  fieldtype=""i4"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  required=""true"" />"
    
    sXML = sXML & Lf & vbTab & Sp & Sp & "<FIELD"
    sXML = sXML & Lf & vbTab & Sp & Sp & "  attrname=""CVERIFY"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  fieldtype=""string"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  required=""true"""
    sXML = sXML & Lf & vbTab & Sp & Sp & "  WIDTH=""32"" />"
    'CLOSE FIELDS
    sXML = sXML & Lf & vbTab & "</FIELDS>"
    'PARAMS
    sXML = sXML & Lf & vbTab & "<PARAMS"
    sXML = sXML & Lf & vbTab & "  DEFAULT_ORDER=""6 4"""
    sXML = sXML & Lf & vbTab & "  LCID=""0"" />"
    'CLOSE METADATA
    sXML = sXML & Lf & "  </METADATA>"
    'OPEN ROWDATA
    sXML = sXML & Lf & Sp & Sp & "<ROWDATA>"
    'ROW
    Dim mRow As Integer
    mRow = 2
    While Cells(mRow, 1) > ""
        sXML = sXML & Lf & vbTab & Oq & "ROW"
        For mCol = 1 To 6
            sXML = sXML & Lf & vbTab & Sp & Sp & Sp & Sp & Trim$(Cells(1, mCol)) & Eq & Q
            sXML = sXML & Trim$(Cells(mRow, mCol)) & Q
        Next
        sXML = sXML & Sp & Sl & Cq
        mRow = mRow + 1
    Wend
    'CLOSE ROWDATA
    sXML = sXML & Lf & Sp & Sp & "</ROWDATA>"
    'CLOSE DATAPACKET
    sXML = sXML & Lf & "</DATAPACKET>"
    
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2 'Specify stream type - we want To save text/string data.
    fsT.Charset = "utf-8" 'Specify charset For the source text data.
    fsT.Open 'Open the stream And write binary data To the object
    fsT.WriteText sXML
    fsT.SaveToFile sOutputFileName, 2 'Save binary data To disk
    
    Close
End Sub
