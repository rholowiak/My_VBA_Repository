Attribute VB_Name = "Module4"
Sub Mail_TabeleZakres()
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Set rng = Range(Selection.Address)
    ''Set rng = Sheets("Sheet1").Range("D1:H4")
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    With OutMail
        .HTMLBody = RangetoHTML(rng)
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
Function RangetoHTML(rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
 
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        Application.CutCopyMode = False
        On Error Resume Next
        On Error GoTo 0
    End With
 
    With TempWB.PublishObjects.Add(SourceType:=xlSourceRange, Filename:=TempFile, Sheet:=TempWB.Sheets(1).Name, Source:=TempWB.Sheets(1).UsedRange.Address, HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
 
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", "align=left x:publishsource=")
  
    TempWB.Close SaveChanges:=False
    Kill TempFile
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
 
'Sub KopiaShipment_Status()
'
'Dim wkbName As String
'
'wkbName = Application.ActiveWorkbook.Name
'
'Application.ActiveSheet.Copy
'
'Application.ActiveWorkbook.SaveAs Filename:="C:\Documents and Settings\LGD\Desktop\" & Left(wkbName, 26) & " Time " & Hour(Time) & "_" & Minute(Time) & ".xls", FileFormat:=-4143
'
'End Sub
Sub Workbook_Mail()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim SigString As String
    Dim Signature As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    SigString = Environ("appdata") & _
                "\Microsoft\Signatures\Rafal2.htm"
    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    On Error Resume Next
    With OutMail
        .To = "renata@lgdisplay.com; michal.stepien@lgdisplay.com; zak.j@lgdisplay.com; kamil.paluszkiewicz@lgdisplay.com; czuba.k@lgdisplay.com; witczak.l@lgdisplay.com;lukasz.nalewalski@lgdisplay.com; maciej.wolak@lgdisplay.com; marcin_pelikan@lgdisplay.com; pytel.m@lgdisplay.com; florczyk.p@lgdisplay.com; tomaszewski.p@lgdisplay.com;pawel.zacharski@lgdisplay.com; czmyr.m@lgdisplay.com; nowakowska.j@lgdisplay.com; yan@lgdisplay.com; kamila.krysa@lgdisplay.com; kamil_zielinski@lgdisplay.com;doll@lgdisplay.com; piotr.kinder@lgdisplay.com; surfingguy@lgdisplay.com;djjang@lgdisplay.com;youngman@lgdisplay.com;marika.marchewka@lgdisplay.com;grzegorz_kunecki@lgdpartner.com;marcin.manijak@lgdisplay.com;malgorzata.turowska@lgdisplay.com;stepniak.s@lgdisplay.com;karisma122@lgdpartner.com;marcin.palinski@lgdisplay.com;i.wrobel@lgdisplay.com;pawel.markowski@lgdisplay.com;k_korzeniowska@lgdpartner.com;agnieszka.kaczynska@lgdisplay.com;woznicka.p@lgdisplay.com;kozuszek.mariusz@lgdisplay.com;sulper@lgdisplay.com"
        .Subject = "Shipment Status " & Format(Now, "dd" & """th""" & " mmmm yyyy" & """ godz. """ & "h:mm")
        .HTMLBody = Signature
        .Attachments.Add ActiveWorkbook.FullName
        ' You can add other files by uncommenting the following line.
        '.Attachments.Add ("C:\test.txt")
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
Function GetBoiler(ByVal sFile As String) As String
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
End Function
 
 
 
Public CELL As Range, CELL2 As Range, CELL3 As Range, CELL5 As Range, myRange As Range, n As Range, CELL6 As Range, CELL7 As Range, CELL8 As Range, CELL9 As Range
Public CELL10 As Variant, CELL11 As Variant, CELL12 As Variant, CELL13 As Variant, CELL14 As Variant, CELL15 As Variant, CELL16 As Variant
Public LastRow As Integer, LastRow2 As Integer, LastRow3 As Integer, LastRow4 As Integer, LastRow5 As Integer, KolumnaHoldReason As Integer
Public LastColumn As Integer, KolumnaDzisProduction As Integer, KolumnaDzis As Integer, KolumnaHold As Integer, KolumnaModel As Integer
Public nResult As Integer, NRofLastList As Integer, SumAi As Integer, h As Integer, i As Integer, c As Integer, R As Long
Public LastColumn5 As Integer, LastRow6 As Integer, LastRow7 As Integer, myRow5 As Integer
Public mySum As Long, mySum2 As Long, mySum3 As Long, mySum4 As Long, mySum5 As Long, mySum6 As Long, mySum7 As Long, mySum8 As Long, mySum9 As Long, mySum10 As Long, mySum11 As Long, mySum12 As Long, mySum13 As Long, mySum14 As Long, mySum15 As Long
Public mojaZmienna1 As Long, mojaZmienna2 As Long, mojaZmienna3 As Long, mojaZmienna4 As Long, myShippingTotal As Long
Public FDofMonth As Date, FDofMonth2 As Date
Public nStart(2) As String, nQuit(2) As String, nAccept(2) As String, Suffix3 As String, firstAddress As String
Public M As Integer, Y As Integer
Sub Shipment_Report2()
    ' Shipment delivery and inventory status Report
    ' Written by Rafal Holowiak 2010-11-09
Const D As Byte = 1
    nStart(2) = "RAFAL HOLOWIAK's PROGRAM"
    nStart(1) = "Hi this is Rafal Holowiak's program" & vbCrLf & "It will help you to create Shipping plan status" & vbCrLf & "Do you need help?"
    nQuit(1) = "Thank You and good luck!"
    nQuit(2) = "PROGRAM QUIT"
    nAccept(1) = "Good choice !" & vbCrLf & "Click OK to proceed"
    nAccept(2) = "PROGRAM START"
    nResult = MsgBox(nStart(1), vbYesNo + vbExclamation, nStart(2))
    If nResult = 6 Then
            MsgBox nAccept(1), vbInformation, nAccept(2)
    Else
        MsgBox nQuit(1), vbCritical, nQuit(2)
        Exit Sub
    End If
FDofMonth2 = Worksheets("Sheet1").Cells(1, 8).Value
Worksheets("Sheet1").Activate
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Range("A1:G2").UnMerge
Range("A1:G1").Cut Destination:=Range("A2:G2")
Rows(1).EntireRow.Delete
Set myRange = Union(Columns(2), Columns(5), Columns(6), Columns(7))
myRange.EntireColumn.Delete
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
For Each CELL In Range(Cells(2, 256), Cells(LastRow, 256))
    Suffix3 = Right((CELL.Offset(0, -255)), 1)
        Select Case Suffix3
            Case "A"
                CELL.Value = 1
            Case "B"
                CELL.Value = 1
            Case "C"
                CELL.Value = 1
            Case "S"
                CELL.Value = 1
            Case "F"
                CELL.Value = 1
            Case "Z"
                CELL.Value = 1
            Case Else
                CELL.Clear
        End Select
Next CELL
i = LastRow - 1
h = Range(Cells(2, 256), Cells(LastRow, 256)).SpecialCells(xlCellTypeConstants).Count
If i <> h Then
Range(Cells(2, 256), Cells(LastRow, 256)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End If
Application.AddCustomList ListArray:=Array("PANASONIC", "TOSHIBA", "SONY", "LGE", "PHILIPS", "ETC-TR", "VESTEL E", "ARCELIK", "LOEWE A", "METZ", "WISTRON O")
NRofLastList = Application.CustomListCount + 1
Worksheets("Sheet1").Cells(1, 1).CurrentRegion.Sort Key1:=Range("B2"), Order1:=xlAscending, Key2:=Range("A2") _
        , Order2:=xlAscending, Header:=xlYes, OrderCustom:=NRofLastList, MatchCase:=True, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
Cells(1, LastColumn + 1).Value = "Shippable"
Cells(1, LastColumn + 2).Value = "|Picked|"
Cells(1, LastColumn + 3).Value = "Waiting for Judgement"
Cells(1, LastColumn + 4).Value = "NG"
Cells(1, LastColumn + 5).Value = "HOLD"
Cells(1, LastColumn + 6).Value = "Under ReWork"
Cells(1, LastColumn + 7).Value = "ReWork req"
With Range(Cells(1, LastColumn + 1), Cells(1, LastColumn + 7))
        .Interior.ColorIndex = 37
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
            With .Font
                    .Name = "Arial"
                    .Size = 9
                    .ColorIndex = xlAutomatic
                    .Bold = True
            End With
End With
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
For Each CELL In Range(Cells(2, 256), Cells(LastRow, 256))
CELL.Value = CELL.Offset(, -255).Value & CELL.Offset(, -254).Value
Next CELL
For Each CELL In Range(Cells(2, 255), Cells(LastRow, 255))
CELL.Value = CELL.Offset(, -254).Value & CELL.Offset(, -253).Value & Cells(1, LastColumn + 1).Value
Next CELL
For Each CELL In Range(Cells(2, 254), Cells(LastRow, 254))
CELL.Value = CELL.Offset(, -253).Value & CELL.Offset(, -252).Value & Cells(1, LastColumn + 2).Value
Next CELL
For Each CELL In Range(Cells(2, 253), Cells(LastRow, 253))
CELL.Value = CELL.Offset(, -252).Value & CELL.Offset(, -251).Value & Cells(1, LastColumn + 3).Value
Next CELL
For Each CELL In Range(Cells(2, 252), Cells(LastRow, 252))
CELL.Value = CELL.Offset(, -251).Value & CELL.Offset(, -250).Value & Cells(1, LastColumn + 4).Value
Next CELL
For Each CELL In Range(Cells(2, 251), Cells(LastRow, 251))
CELL.Value = CELL.Offset(, -250).Value & CELL.Offset(, -249).Value & Cells(1, LastColumn + 5).Value
Next CELL
For Each CELL In Range(Cells(2, 250), Cells(LastRow, 250))
CELL.Value = CELL.Offset(, -249).Value & CELL.Offset(, -248).Value & Cells(1, LastColumn + 6).Value
Next CELL
For Each CELL In Range(Cells(2, 249), Cells(LastRow, 249))
CELL.Value = CELL.Offset(, -248).Value & CELL.Offset(, -247).Value & Cells(1, LastColumn + 7).Value
Next CELL
LastRow3 = Worksheets("PalletList").Cells(Rows.Count, 1).End(xlUp).Row
For Each CELL2 In Worksheets("PalletList").Range(Worksheets("PalletList").Cells(2, 255), Worksheets("PalletList").Cells(LastRow3, 255))
    If ((CELL2.Offset(, -247).Value) <> 0) And (((CELL2.Offset(, -249).Value = "REWORK") = True) Or (((CELL2.Offset(, -249).Value = "INT MFG") = True))) Then
        CELL2.Value = "Under ReWork"
    ElseIf ((CELL2.Offset(, -247).Value) <> 0) And (((CELL2.Offset(, -249).Value <> "REWORK") = True) And (((CELL2.Offset(, -249).Value <> "INT MFG") = True))) Then
        CELL2.Value = "ReWork req"
    ElseIf ((CELL2.Offset(, -248).Value) <> 0) Then
        CELL2.Value = "HOLD"
    ElseIf ((CELL2.Offset(, -249).Value) = "STG") Then
        CELL2.Value = "|Picked|"
    ElseIf ((CELL2.Offset(, -240).Value) = "RS") Then
        CELL2.Value = "Waiting for Judgement"
    ElseIf ((CELL2.Offset(, -240).Value) = "NG") Then
        CELL2.Value = "NG"
    Else
    CELL2.Value = "Shippable"
    End If
Next CELL2
For Each CELL2 In Worksheets("PalletList").Range(Worksheets("PalletList").Cells(2, 256), Worksheets("PalletList").Cells(LastRow3, 256))
        CELL2.Value = CELL2.Offset(, -254).Value & CELL2.Offset(, -244).Value & CELL2.Offset(, -1).Value
Next CELL2
For Each CELL2 In Worksheets("PalletList").Range(Worksheets("PalletList").Cells(2, 254), Worksheets("PalletList").Cells(LastRow3, 254))
        CELL2.Value = CELL2.Offset(, -252).Value & CELL2.Offset(, 1).Value
Next CELL2
For Each CELL In Range(Cells(2, LastColumn + 1), Cells(LastRow, LastColumn + 1))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, 256 - LastColumn - 2), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(2, LastColumn + 2), Cells(LastRow, LastColumn + 2))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 4)), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(2, LastColumn + 3), Cells(LastRow, LastColumn + 3))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 6)), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(2, LastColumn + 4), Cells(LastRow, LastColumn + 4))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 8)), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(2, LastColumn + 5), Cells(LastRow, LastColumn + 5))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 10)), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(2, LastColumn + 6), Cells(LastRow, LastColumn + 6))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 12)), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(2, LastColumn + 7), Cells(LastRow, LastColumn + 7))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 14)), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
With Range(Cells(2, LastColumn + 1), Cells(LastRow, LastColumn + 7))
      .Borders(xlDiagonalDown).LineStyle = xlNone
      .Borders(xlDiagonalUp).LineStyle = xlNone
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous
End With
With Worksheets("Sheet1").Range(Cells(1, 4), Cells(1, LastColumn))
    Set CELL = .Find("Total", LookIn:=xlValues)
    If Not CELL Is Nothing Then
            myShippingTotal = Application.WorksheetFunction.Sum(Range(CELL.Offset(1, 0), CELL.Offset(LastRow - 1, 0)))
    Else
            myShippingTotal = "no data"
    End If
End With
With Worksheets("Sheet1").Range(Cells(1, 4), Cells(1, LastColumn))
    Set CELL = .Find("Total", LookIn:=xlValues)
    If Not CELL Is Nothing Then
        Do
            Range(CELL.Offset(1, 0), CELL.Offset(LastRow - 1, 0)).Clear
            Range(CELL.Offset(1, 0), CELL.Offset(LastRow - 1, 0)).Interior.ColorIndex = 53
            CELL.Value = "\"
            Set CELL = .FindNext(CELL)
        Loop While Not CELL Is Nothing
    End If
End With
Range(Cells(1, 1), Cells(LastRow, 1)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Cells(LastRow + 4, 1), Unique:=True
LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(LastRow + 4, 1), Cells(LastRow2, 1)).Sort Key1:=Range("A301"), Order1:=xlAscending, Header:=xlGuess _
        , OrderCustom:=1, MatchCase:=True, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
Cells(LastRow2 + 1, 1).Value = "TOTAL"
LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(1, 2), Cells(1, LastColumn)).Copy Range(Cells(LastRow + 4, 2), Cells(LastRow + 4, 2))
Cells(LastRow + 3, 1) = "Calculation by Model Name"
Cells(LastRow + 3, 2) = "All"
Cells(LastRow + 3, 3) = "All"
Range(Cells(1, 1), Cells(1, 3)).Copy
With Range(Cells(LastRow + 3, 1), Cells(LastRow + 3, 3))
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .Font.ColorIndex = 0
End With
Application.CutCopyMode = False
For Each CELL In Range(Cells(LastRow + 5, 4), Cells(LastRow2, LastColumn))
    c = CELL.Column
    CELL.Value = Application.WorksheetFunction. _
        SumIf(Range(Cells(2, 1), Cells(LastRow, 1)), CELL.Offset(, -c + 1), Range(Cells(2, c), Cells(LastRow, c)))
Next CELL
For Each CELL In Range(Cells(LastRow + 1, 4), Cells(LastRow + 1, LastColumn))
    mySum = Application.WorksheetFunction.Sum(Range(CELL.Offset(-LastRow + 1, 0), CELL.Offset(-1, 0)))
        With CELL
            .Value = mySum
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThick
        End With
Next CELL
Range(Cells(2, 2), Cells(2, LastColumn)).Copy
With Range(Cells(LastRow + 5, 2), Cells(LastRow2, LastColumn))
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .Font.ColorIndex = 0
End With
Application.CutCopyMode = False
Cells(LastRow + 4, LastColumn + 1).Value = "Shippable"
Cells(LastRow + 4, LastColumn + 2).Value = "|Picked|"
Cells(LastRow + 4, LastColumn + 3).Value = "Waiting for Judgement"
Cells(LastRow + 4, LastColumn + 4).Value = "NG"
Cells(LastRow + 4, LastColumn + 5).Value = "HOLD"
Cells(LastRow + 4, LastColumn + 6).Value = "Under ReWork"
Cells(LastRow + 4, LastColumn + 7).Value = "ReWork req"
Cells(LastRow + 4, LastColumn + 8).Value = "CUMM.PLAN TO SHIP till today"
Cells(LastRow + 4, LastColumn + 9).Value = "CUMM.DISPATCHED till today 7am"
Cells(LastRow + 4, LastColumn + 10).Value = "REMAIN To Complete TARGET"
Cells(LastRow + 4, LastColumn + 11).Value = "MISSING for Delivery D"
Cells(LastRow + 4, LastColumn + 12).Value = "MISSING for Delivery D+1"
Cells(LastRow + 4, LastColumn + 13).Value = "MISSING for Delivery D+2"
Cells(LastRow + 4, LastColumn + 14).Value = "MISSING for Delivery D+3"
Cells(LastRow + 4, LastColumn + 15).Value = "MISSING for Delivery D+4"
With Range(Cells(LastRow + 4, LastColumn + 1), Cells(LastRow + 4, LastColumn + 15))
      .Interior.ColorIndex = 37
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = True
End With
Range(Cells(1, 4), Cells(1, LastColumn)).Copy Cells(65536, 4)
Range(Cells(65536, 4), Cells(65536, LastColumn)).NumberFormat = "[$-415]dmmm;@"
h = 0
For Each CELL In Range(Cells(65536, 4), Cells(65536, LastColumn))
    If CELL.Value <> "\" Then
         CELL.Value = FDofMonth2 + h
         h = h + 1
    End If
Next CELL
KolumnaDzis = Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Column
With Range(Cells(2, KolumnaDzis), Cells(LastRow, KolumnaDzis))
    .Interior.ColorIndex = 0
End With
With Range(Cells(LastRow + 5, KolumnaDzis), Cells(LastRow2, KolumnaDzis))
    .Interior.ColorIndex = 0
End With
For Each CELL In Range(Cells(LastRow + 5, 255), Cells(LastRow2, 255))
CELL.Value = CELL.Offset(, -254).Value & Cells(LastRow + 4, LastColumn + 1).Value
Next CELL
For Each CELL In Range(Cells(LastRow + 5, 254), Cells(LastRow2, 254))
CELL.Value = CELL.Offset(, -253).Value & Cells(LastRow + 4, LastColumn + 2).Value
Next CELL
For Each CELL In Range(Cells(LastRow + 5, 253), Cells(LastRow2, 253))
CELL.Value = CELL.Offset(, -252).Value & Cells(LastRow + 4, LastColumn + 3).Value
Next CELL
For Each CELL In Range(Cells(LastRow + 5, 252), Cells(LastRow2, 252))
CELL.Value = CELL.Offset(, -251).Value & Cells(LastRow + 4, LastColumn + 4).Value
Next CELL
For Each CELL In Range(Cells(LastRow + 5, 251), Cells(LastRow2, 251))
CELL.Value = CELL.Offset(, -250).Value & Cells(LastRow + 4, LastColumn + 5).Value
Next CELL
For Each CELL In Range(Cells(LastRow + 5, 250), Cells(LastRow2, 250))
CELL.Value = CELL.Offset(, -249).Value & Cells(LastRow + 4, LastColumn + 6).Value
Next CELL
For Each CELL In Range(Cells(LastRow + 5, 249), Cells(LastRow2, 249))
CELL.Value = CELL.Offset(, -248).Value & Cells(LastRow + 4, LastColumn + 7).Value
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 1), Cells(LastRow2, LastColumn + 1))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 2), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 2), Cells(LastRow2, LastColumn + 2))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 4), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 3), Cells(LastRow2, LastColumn + 3))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 6), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 4), Cells(LastRow2, LastColumn + 4))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 8), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 5), Cells(LastRow2, LastColumn + 5))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 10), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 6), Cells(LastRow2, LastColumn + 6))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 12), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 7), Cells(LastRow2, LastColumn + 7))
   
    With Worksheets("PalletList")
        CELL.Value = Application.WorksheetFunction. _
            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 14), .Range(.Cells(2, 3), .Cells(LastRow3, 3)))
    End With
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 8), Cells(LastRow2, LastColumn + 8))
    CELL.Value = Application.WorksheetFunction.Sum(Range(Cells(CELL.Row, 4), Cells(CELL.Row, KolumnaDzis)))
Next CELL
LastRow5 = Worksheets("Shipment Status").Cells(Rows.Count, 1).End(xlUp).Row
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 9), Cells(LastRow2, LastColumn + 9))
    CELL.Value = Application.WorksheetFunction.SumIf(Worksheets("Shipment Status") _
                    .Range(Worksheets("Shipment Status").Cells(2, 5), Worksheets("Shipment Status") _
                        .Cells(LastRow5, 5)), CELL.Offset(0, -LastColumn - 8), Worksheets("Shipment Status") _
                            .Range(Worksheets("Shipment Status").Cells(2, 23), Worksheets("Shipment Status").Cells(LastRow5, 23)))
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 10), Cells(LastRow2, LastColumn + 10))
    CELL.Value = (CELL.Offset(0, -1).Value - CELL.Offset(0, -2).Value)
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 11), Cells(LastRow2, LastColumn + 11))
    CELL.Value = (CELL.Offset(0, -1).Value + CELL.Offset(0, -10).Value)
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 12), Cells(LastRow2, LastColumn + 12))
    mojaZmienna1 = Application.WorksheetFunction.Sum(Range(Cells(CELL.Row, 4), Cells(CELL.Row, KolumnaDzis + 1)))
    mojaZmienna2 = CELL.Offset(0, -3).Value
    mojaZmienna3 = mojaZmienna2 - mojaZmienna1
    mojaZmienna4 = mojaZmienna3 + CELL.Offset(0, -11).Value
    CELL.Value = mojaZmienna4
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 13), Cells(LastRow2, LastColumn + 13))
    mojaZmienna1 = Application.WorksheetFunction.Sum(Range(Cells(CELL.Row, 4), Cells(CELL.Row, KolumnaDzis + 2)))
    mojaZmienna2 = CELL.Offset(0, -4).Value
    mojaZmienna3 = mojaZmienna2 - mojaZmienna1
    mojaZmienna4 = mojaZmienna3 + CELL.Offset(0, -12).Value
    CELL.Value = mojaZmienna4
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 14), Cells(LastRow2, LastColumn + 14))
    mojaZmienna1 = Application.WorksheetFunction.Sum(Range(Cells(CELL.Row, 4), Cells(CELL.Row, KolumnaDzis + 3)))
    mojaZmienna2 = CELL.Offset(0, -5).Value
    mojaZmienna3 = mojaZmienna2 - mojaZmienna1
    mojaZmienna4 = mojaZmienna3 + CELL.Offset(0, -13).Value
    CELL.Value = mojaZmienna4
Next CELL
For Each CELL In Range(Cells(LastRow + 5, LastColumn + 15), Cells(LastRow2, LastColumn + 15))
    mojaZmienna1 = Application.WorksheetFunction.Sum(Range(Cells(CELL.Row, 4), Cells(CELL.Row, KolumnaDzis + 4)))
    mojaZmienna2 = CELL.Offset(0, -6).Value
    mojaZmienna3 = mojaZmienna2 - mojaZmienna1
    mojaZmienna4 = mojaZmienna3 + CELL.Offset(0, -14).Value
    CELL.Value = mojaZmienna4
Next CELL
For Each CELL In Range(Cells(LastRow2, 2), Cells(LastRow2, LastColumn + 15))
    mySum = Application.WorksheetFunction.Sum(Range(CELL.Offset(LastRow - LastRow2 + 5, 0), CELL.Offset(-1, 0)))
        With CELL
            .Value = mySum
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThick
            .Interior.ColorIndex = 37
            .NumberFormat = "#,##"
        End With
Next CELL
With Range(Cells(LastRow + 4, 1), Cells(LastRow2, LastColumn + 15))
        .Borders(xlEdgeLeft).LineStyle = xlDouble
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        .Borders(xlEdgeTop).LineStyle = xlDouble
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).LineStyle = xlDouble
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlEdgeRight).LineStyle = xlDouble
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeRight).ColorIndex = xlAutomatic
        .Borders(xlInsideVertical).LineStyle = xlDash
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideVertical).ColorIndex = xlAutomatic
        .Borders(xlInsideHorizontal).LineStyle = xlDash
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
End With
Range(Cells(LastRow + 4, LastColumn + 8), Cells(LastRow2 - 1, LastColumn + 8)).Interior.ColorIndex = 35
Range(Cells(LastRow + 4, LastColumn + 9), Cells(LastRow2 - 1, LastColumn + 9)).Interior.ColorIndex = 50
With Range(Cells(LastRow + 4, LastColumn + 10), Cells(LastRow2 - 1, LastColumn + 10))
    .Interior.ColorIndex = 35
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    With .FormatConditions(1).Font
                .Bold = True
                .Italic = False
                .ColorIndex = 3
    End With
End With
Range(Cells(LastRow + 5, LastColumn + 1), Cells(LastRow2 - 1, LastColumn + 7)).Interior.ColorIndex = 36
With Range(Cells(LastRow + 5, LastColumn + 11), Cells(LastRow2 - 1, LastColumn + 15))
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
        With .FormatConditions(1)
                .Interior.ColorIndex = 15
                .Font.Bold = True
                .Font.Italic = False
                .Font.ColorIndex = 3
        End With
    .FormatConditions(2).Font.ColorIndex = 2
End With
With Range(Cells(LastRow + 4, LastColumn + 1), Cells(LastRow + 4, LastColumn + 15)).Font
        .Name = "Arial"
        .Size = 9
        .ColorIndex = xlAutomatic
End With
Cells(LastRow + 1, LastColumn + 9).Value = "Shipment status"
Cells(LastRow + 1, LastColumn + 11).Value = "Shipment Total"
Cells(LastRow + 2, LastColumn + 11).Value = "Missing Total"
With Range(Cells(LastRow + 1, LastColumn + 8), Cells(LastRow + 2, LastColumn + 13))
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Interior.ColorIndex = 4
        .Interior.Pattern = xlSolid
        .Font.Bold = True
        .NumberFormat = "#,##0"
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
End With
Cells(LastRow + 1, LastColumn + 2).Value = "Inventory status"
Cells(LastRow + 1, LastColumn + 4).Value = "Total FGS"
With Range(Cells(LastRow + 1, LastColumn + 1), Cells(LastRow + 2, LastColumn + 7))
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .Interior.ColorIndex = 6
        .Interior.Pattern = xlSolid
        .Font.Bold = True
        .NumberFormat = "#,##0"
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
End With
Cells(LastRow + 2, LastColumn + 1).Value = Cells(LastRow2, LastColumn + 1).Value
Cells(LastRow + 2, LastColumn + 2).Value = Cells(LastRow2, LastColumn + 2).Value
Cells(LastRow + 2, LastColumn + 3).Value = Cells(LastRow2, LastColumn + 3).Value
Cells(LastRow + 2, LastColumn + 4).Value = Cells(LastRow2, LastColumn + 4).Value
Cells(LastRow + 2, LastColumn + 5).Value = Cells(LastRow2, LastColumn + 5).Value
Cells(LastRow + 2, LastColumn + 6).Value = Cells(LastRow2, LastColumn + 6).Value
Cells(LastRow + 2, LastColumn + 7).Value = Cells(LastRow2, LastColumn + 7).Value
Cells(LastRow + 2, LastColumn + 8).Value = Cells(LastRow2, LastColumn + 8).Value
Cells(LastRow + 2, LastColumn + 9).Value = Cells(LastRow2, LastColumn + 9).Value
Cells(LastRow + 1, LastColumn + 5).Value = Application.WorksheetFunction.Sum(Range(Cells(LastRow2, LastColumn + 1), Cells(LastRow2, LastColumn + 7)))
Cells(LastRow + 1, LastColumn + 12).Value = myShippingTotal
Cells(LastRow + 2, LastColumn + 12).Value = (Cells(LastRow + 2, LastColumn + 9).Value) - (Cells(LastRow + 1, LastColumn + 12).Value)
With Cells(LastRow + 2, LastColumn + 12)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        With .FormatConditions(1).Font
                .ColorIndex = 3
        End With
End With
With Range(Cells(2, 3), Cells(LastRow, 3))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 21
End With
Range(Cells(LastRow + 1, 1), Cells(LastRow + 4, 1)).EntireRow.AutoFit
Union(Range(Cells(2, 1), Cells(LastRow, 1)), Range(Cells(LastRow + 5, 1), Cells(LastRow2 - 1, 1))).EntireRow.RowHeight = 11.25
Rows(LastRow2).RowHeight = 15
Range(Cells(1, 4), Cells(1, LastColumn)).EntireColumn.ColumnWidth = 6.57
Range(Cells(1, LastColumn + 8), Cells(1, LastColumn + 10)).EntireColumn.ColumnWidth = 10
Cells(1, LastColumn + 1).EntireColumn.ColumnWidth = 8.43
Range(Cells(1, LastColumn + 2), Cells(1, LastColumn + 7)).EntireColumn.ColumnWidth = 7.29
Range(Cells(1, LastColumn + 11), Cells(1, LastColumn + 13)).EntireColumn.ColumnWidth = 7.29
Columns(LastColumn + 3).ColumnWidth = 9.86
Columns(LastColumn + 12).ColumnWidth = 7.86
With Worksheets("Sheet1").Range(Cells(1, 4), Cells(1, LastColumn))
    Set CELL = .Find("\", LookIn:=xlValues)
    If Not CELL Is Nothing Then
        Do
            CELL.EntireColumn.ColumnWidth = 0.92
            CELL.Value = "|"
            Set CELL = .FindNext(CELL)
        Loop While Not CELL Is Nothing
    End If
End With
Worksheets("HOLD").Activate
Worksheets("HOLD").Cells(1, 1).Select
With Worksheets("HOLD")
    LastRow4 = .Cells(Rows.Count, 1).End(xlUp).Row
    KolumnaHold = .Cells.Find(What:="Type", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Column
    KolumnaModel = .Cells.Find(What:="Model", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Column
    KolumnaHoldReason = .Cells.Find(What:="Holding Reason", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Column
    KolumnaCustomer = .Cells.Find(What:="Customer", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Column
    .Cells(1, 1).CurrentRegion. _
        Sort Key1:=.Columns(KolumnaHold), Order1:=xlAscending, Key2:=.Columns(KolumnaModel) _
            , Order2:=xlAscending, Key3:=.Columns(KolumnaHoldReason), Order3:=xlAscending, Header:= _
            xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
            xlSortNormal
    For Each CELL3 In Range(.Cells(2, KolumnaHold), .Cells(LastRow4, KolumnaHold))
        If CELL3.Value <> "HOLD" Then CELL3.Clear
    Next CELL3
    i = LastRow4 - 1
    h = Range(.Cells(2, KolumnaHold), .Cells(LastRow4, KolumnaHold)).SpecialCells(xlCellTypeConstants).Count
    If i <> h Then
            Range(.Cells(2, KolumnaHold), .Cells(LastRow4, KolumnaHold)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    End If
    LastColumn5 = .Cells(1, Columns.Count).End(xlToLeft).Column
    LastRow6 = .Cells(Rows.Count, 1).End(xlUp).Row
    For Each CELL5 In Range(.Cells(2, 100), .Cells(LastRow6, 100))
        myRow5 = CELL5.Row
        Set CELL8 = .Cells(myRow5, KolumnaModel)
        Set CELL9 = .Cells(myRow5, KolumnaHoldReason)
        Set CELL6 = .Cells(myRow5 - 1, KolumnaModel)
        Set CELL7 = .Cells(myRow5 - 1, KolumnaHoldReason)
        If CELL8 = CELL6 Then
            CELL5.Value = CELL5.Offset(-1, 0).Value & " / " & CELL9.Value
        Else
            CELL5.Value = CELL9.Value
        End If
    Next CELL5
    For Each CELL5 In Range(.Cells(2, 101), .Cells(LastRow6, 101))
        myRow5 = CELL5.Row
        Set CELL8 = .Cells(myRow5, KolumnaModel)
            For i = 1 To 7
                Set CELL9 = .Cells(myRow5 + i, KolumnaModel)
                    If CELL8.Value <> CELL9.Value Then
                        CELL5.Value = CELL5.Offset(i - 1, -1).Value
                        Exit For
                    Else
                        CELL5.Value = CELL5.Offset(i, -1).Value
                    End If
            Next i
    Next CELL5
    For Each CELL5 In Range(.Cells(2, 99), .Cells(LastRow6, 99))
        myRow5 = CELL5.Row
        Set CELL10 = .Cells(myRow5, KolumnaModel)
        Set CELL11 = .Cells(myRow5, KolumnaCustomer)
        Set CELL12 = .Cells(myRow5, KolumnaHold)
        CELL5.Value = CELL10.Value & CELL11.Value & CELL12.Value
    Next CELL5
End With
Worksheets("Sheet1").Activate
Worksheets("Sheet1").Cells(1, 1).Select
For Each CELL In Range(Cells(2, LastColumn + 8), Cells(LastRow, LastColumn + 8))
        Set CELL2 = Cells(CELL.Row, 251)
            On Error Resume Next
            CELL.Value = Application.WorksheetFunction.VLookup(CELL2.Value, Range(Worksheets("HOLD").Cells(2, 99), Worksheets("HOLD").Cells(LastRow6, 101)), 3, False)
Next CELL
Range(Cells(1, KolumnaDzis + 5), Cells(1, KolumnaDzis + 8)).EntireColumn.Insert
Range(Cells(1, KolumnaDzis), Cells(1, KolumnaDzis + 3)).Copy Range(Cells(1, KolumnaDzis + 5), Cells(1, KolumnaDzis + 8))
CELL13 = Worksheets("Sheet1").Cells(1, KolumnaDzis + 5).Value
CELL14 = Worksheets("Sheet1").Cells(1, KolumnaDzis + 6).Value
CELL15 = Worksheets("Sheet1").Cells(1, KolumnaDzis + 7).Value
CELL16 = Worksheets("Sheet1").Cells(1, KolumnaDzis + 8).Value
Cells(1, KolumnaDzis + 5).Value = "Production " & CELL13
Cells(1, KolumnaDzis + 6).Value = "Production " & CELL14
Cells(1, KolumnaDzis + 7).Value = "Production " & CELL15
Cells(1, KolumnaDzis + 8).Value = "Production " & CELL16
Cells(LastRow + 4, KolumnaDzis + 5).Value = "Production " & CELL13
Cells(LastRow + 4, KolumnaDzis + 6).Value = "Production " & CELL14
Cells(LastRow + 4, KolumnaDzis + 7).Value = "Production " & CELL15
Cells(LastRow + 4, KolumnaDzis + 8).Value = "Production " & CELL16
Range(Cells(1, KolumnaDzis + 5), Cells(1, KolumnaDzis + 8)).EntireColumn.ColumnWidth = 8.29
With Range(Cells(1, KolumnaDzis + 5), Cells(LastRow, KolumnaDzis + 8))
        .Interior.Color = RGB(141, 180, 226)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
End With
With Range(Cells(LastRow + 4, KolumnaDzis + 5), Cells(LastRow2 - 1, KolumnaDzis + 8))
        .Interior.Color = RGB(141, 180, 226)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
End With
Worksheets(5).Activate
KolumnaDzisProduction = Worksheets(5).Cells.Find(What:=Date, After:=Cells(5, Columns.Count).End(xlToLeft), LookIn:=xlFormulas, LookAt:=xlPart, _
    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False).Column
Worksheets("Sheet1").Activate
Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, KolumnaDzis + 5), Worksheets("Sheet1").Cells(LastRow, KolumnaDzis + 5))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction), Worksheets(5).Cells(300, KolumnaDzisProduction)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1), Worksheets(5).Range(Worksheets(5).Cells(7, 10), Worksheets(5).Cells(300, 10)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, KolumnaDzis + 6), Worksheets("Sheet1").Cells(LastRow, KolumnaDzis + 6))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction + 1), Worksheets(5).Cells(300, KolumnaDzisProduction + 1)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1), Worksheets(5).Range(Worksheets(5).Cells(7, 10), Worksheets(5).Cells(300, 10)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, KolumnaDzis + 7), Worksheets("Sheet1").Cells(LastRow, KolumnaDzis + 7))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction + 2), Worksheets(5).Cells(300, KolumnaDzisProduction + 2)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1), Worksheets(5).Range(Worksheets(5).Cells(7, 10), Worksheets(5).Cells(300, 10)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, KolumnaDzis + 8), Worksheets("Sheet1").Cells(LastRow, KolumnaDzis + 8))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction + 3), Worksheets(5).Cells(300, KolumnaDzisProduction + 3)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1), Worksheets(5).Range(Worksheets(5).Cells(7, 10), Worksheets(5).Cells(300, 10)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(LastRow + 5, KolumnaDzis + 5), Worksheets("Sheet1").Cells(LastRow2 - 1, KolumnaDzis + 5))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction), Worksheets(5).Cells(300, KolumnaDzisProduction)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(LastRow + 5, KolumnaDzis + 6), Worksheets("Sheet1").Cells(LastRow2 - 1, KolumnaDzis + 6))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction + 1), Worksheets(5).Cells(300, KolumnaDzisProduction + 1)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(LastRow + 5, KolumnaDzis + 7), Worksheets("Sheet1").Cells(LastRow2 - 1, KolumnaDzis + 7))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction + 2), Worksheets(5).Cells(300, KolumnaDzisProduction + 2)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(LastRow + 5, KolumnaDzis + 8), Worksheets("Sheet1").Cells(LastRow2 - 1, KolumnaDzis + 8))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction + 3), Worksheets(5).Cells(300, KolumnaDzisProduction + 3)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
On Error GoTo 0
    If KolumnaDzis = 4 Then
    Range(Columns(KolumnaDzis + 9), Columns(LastColumn + 4)).EntireColumn.Hidden = True
    Else:
    Range(Columns(4), Columns(KolumnaDzis - 1)).EntireColumn.Hidden = True
    Range(Columns(KolumnaDzis + 9), Columns(LastColumn + 4)).EntireColumn.Hidden = True
    End If
Set myRange = Union(Range(Cells(2, 247), Cells(LastRow, 247)), Range(Cells(LastRow + 5, 247), Cells(LastRow2, 247)))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.Sum(Range(Cells(CELL.Row, KolumnaDzis), Cells(CELL.Row, KolumnaDzis + 4)))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
myRange.SpecialCells(xlCellTypeBlanks).EntireRow.Hidden = True
ActiveWindow.Zoom = 95
ActiveWindow.DisplayZeros = False
With Worksheets("Sheet1")
        .Columns(1).Font.Bold = True
        .Columns(1).EntireColumn.AutoFit
        .Columns(2).EntireColumn.AutoFit
        .Columns(3).EntireColumn.AutoFit
End With
Worksheets("Sheet1").Cells(1, 1).Select
For Each CELL In Range(Cells(1, KolumnaDzis + 5), Cells(1, KolumnaDzis + 8))
    If CELL.Value = "Production |" Then
         CELL.EntireColumn.Hidden = True
    End If
Next CELL
Shipment_Summary
Worksheets("Sheet1").Activate
Worksheets("Sheet1").Cells(1, 1).Select
Kolor_Lampki_1
Set myRange = Range(Cells(LastRow + 5, LastColumn + 15), Cells(LastRow2 - 1, LastColumn + 19))
For Each CELL In myRange
i = CELL.Row
h = CELL.Column
    For c = 1 To 4
        If CELL.Value = Cells(i, h + c).Value Then Cells(i, h + c).Interior.ColorIndex = 2
        If CELL.Value = Cells(i, h + c).Value Then Cells(i, h + c).ClearContents
    Next c
Next CELL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''Kolor_Lampki_2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.DisplayAlerts = False
Worksheets(Worksheets.Count - 2).Delete
Worksheets(Worksheets.Count - 2).Delete
Worksheets(Worksheets.Count - 2).Delete
Worksheets(Worksheets.Count - 2).Delete
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
MsgBox "CREATION COMPLETED", vbInformation, "Program Status"
End Sub
Private Sub Shipment_Summary()
Application.ActiveWorkbook.Worksheets("Sheet1").Copy After:=Worksheets(Worksheets.Count)
Worksheets(Worksheets.Count).Activate
Range(Cells(LastRow + 3, 1), Cells(LastRow2, 1)).EntireRow.Delete
Set myRange = Range(Cells(2, 1), Cells(LastRow, 1))
For Each CELL In myRange
i = CELL.Row
    If CELL.Value = Cells(i + 1, 1).Value And Cells(i, 2).Value = Cells(i + 1, 2).Value Then
            Union(Range(Cells(i + 1, KolumnaDzis + 5), Cells(i + 1, KolumnaDzis + 8)), Range(Cells(i + 1, LastColumn + 5), Cells(i + 1, LastColumn + 11))).ClearContents
    End If
Next CELL
Application.ActiveWorkbook.Worksheets(Worksheets.Count).Copy After:=Worksheets(Worksheets.Count)
Worksheets(Worksheets.Count - 1).Activate
Columns(1).Delete
Range(Cells(1, 1), Cells(LastRow, 1)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Cells(LastRow + 4, 1), Unique:=True
Range(Cells(1, 1), Cells(1, LastColumn + 11)).Copy Range(Cells(LastRow + 4, 1), Cells(LastRow + 4, LastColumn + 11))
LastRow7 = Cells(Rows.Count, 1).End(xlUp).Row

Set myRange = Range(Cells(LastRow + 5, 3), Cells(LastRow7, KolumnaDzis - 1))
For Each CELL In myRange
   
   mySum = Application.WorksheetFunction.SumIf(Range(Cells(2, 1), Cells(LastRow, 1)), Cells(CELL.Row, 1), Range(Cells(2, CELL.Column), Cells(LastRow, CELL.Column)))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, KolumnaDzis + 4), Cells(LastRow7, KolumnaDzis + 4))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, KolumnaDzis + 4), Cells(LastRow, KolumnaDzis + 4)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, KolumnaDzis + 5), Cells(LastRow7, KolumnaDzis + 5))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, KolumnaDzis + 5), Cells(LastRow, KolumnaDzis + 5)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, KolumnaDzis + 6), Cells(LastRow7, KolumnaDzis + 6))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, KolumnaDzis + 6), Cells(LastRow, KolumnaDzis + 6)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, KolumnaDzis + 7), Cells(LastRow7, KolumnaDzis + 7))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, KolumnaDzis + 7), Cells(LastRow, KolumnaDzis + 7)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Range(Cells(2, 1), Cells(2, LastColumn + 10)).Copy
With Range(Cells(LastRow + 5, 1), Cells(LastRow7, LastColumn + 10))
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    .Font.ColorIndex = 0
End With
Application.CutCopyMode = False
Set myRange = Range(Cells(LastRow + 5, LastColumn + 4), Cells(LastRow7, LastColumn + 4))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 4), Cells(LastRow, LastColumn + 4)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 5), Cells(LastRow7, LastColumn + 5))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 5), Cells(LastRow, LastColumn + 5)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 6), Cells(LastRow7, LastColumn + 6))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 6), Cells(LastRow, LastColumn + 6)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 7), Cells(LastRow7, LastColumn + 7))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 7), Cells(LastRow, LastColumn + 7)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 8), Cells(LastRow7, LastColumn + 8))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 8), Cells(LastRow, LastColumn + 8)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 9), Cells(LastRow7, LastColumn + 9))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 9), Cells(LastRow, LastColumn + 9)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 10), Cells(LastRow7, LastColumn + 10))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 10), Cells(LastRow, LastColumn + 10)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Range(Rows(1), Rows(LastRow + 3)).EntireRow.Delete
Range(Columns(KolumnaDzis), Columns(KolumnaDzis + 3)).EntireColumn.Hidden = True
Range(Columns(LastColumn + 4), Columns(LastColumn + 5)).EntireColumn.Hidden = True
Columns(LastColumn + 6).Insert
Columns(LastColumn + 6).Insert
Columns(LastColumn + 6).Insert
Columns(LastColumn + 6).Insert
Cells(1, LastColumn + 6).Value = "PLAN"
Cells(1, LastColumn + 7).Value = "DISPATCHED"
Cells(1, LastColumn + 8).Value = "REMAIN"
Cells(1, LastColumn + 9).Value = "Shipppable"
Columns(LastColumn + 6).ColumnWidth = 10
Columns(LastColumn + 7).ColumnWidth = 11
Columns(LastColumn + 8).ColumnWidth = 11
Columns(LastColumn + 9).ColumnWidth = 11
LastRow7 = Cells(Rows.Count, 1).End(xlUp).Row

Set myRange = Range(Cells(2, LastColumn + 6), Cells(LastRow7, LastColumn + 6))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.Sum(Range(Cells(CELL.Row, 3), Cells(CELL.Row, KolumnaDzis - 1)))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(2, LastColumn + 7), Cells(LastRow7, LastColumn + 7))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIf(Worksheets("Shipment Status").Range(Worksheets("Shipment Status").Cells(2, 11), Worksheets("Shipment Status").Cells(LastRow5, 11)), Cells(CELL.Row, 1), Worksheets("Shipment Status").Range(Worksheets("Shipment Status").Cells(2, 23), Worksheets("Shipment Status").Cells(LastRow5, 23)))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
For Each CELL In Range(Cells(2, LastColumn + 8), Cells(LastRow7, LastColumn + 8))
    CELL.Value = (CELL.Offset(0, -1).Value - CELL.Offset(0, -2).Value)
Next CELL

With Range(Cells(2, LastColumn + 8), Cells(LastRow7, LastColumn + 8))
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
        With .FormatConditions(1)
                .Interior.Color = RGB(234, 234, 234)
                .Font.Bold = True
                .Font.Italic = False
                .Font.ColorIndex = 3
        End With
    .FormatConditions(2).Font.Bold = True
End With
For Each CELL In Range(Cells(2, LastColumn + 9), Cells(LastRow7, LastColumn + 9))
    CELL.Value = CELL.Offset(0, -5).Value
Next CELL

Columns(2).EntireColumn.Delete
Worksheets(Worksheets.Count).Activate
Columns(1).Insert
Cells(1, 1).Value = "Inch"
Columns(2).Copy
With Columns(1)
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .Font.ColorIndex = 0
End With
Application.CutCopyMode = False
Set myRange = Range(Cells(2, 1), Cells(LastRow, 1))
For Each CELL In myRange
CELL.Value = Mid(Cells(CELL.Row, 2), 3, 2)
Next CELL
Columns(2).EntireColumn.Delete
Columns(3).EntireColumn.Delete
Range(Cells(1, 1), Cells(LastRow, 2)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Cells(LastRow + 4, 1), Unique:=True
Range(Cells(1, 1), Cells(1, LastColumn + 11)).Copy Range(Cells(LastRow + 4, 1), Cells(LastRow + 4, LastColumn + 11))
LastRow7 = Cells(Rows.Count, 1).End(xlUp).Row
ActiveWorkbook.Worksheets(Worksheets.Count).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(Worksheets.Count).Sort.SortFields.Add Key _
        :=Range(Cells(LastRow + 5, 1), Cells(LastRow7, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
With ActiveWorkbook.Worksheets(Worksheets.Count).Sort
        .SetRange Range(Cells(LastRow + 4, 1), Cells(LastRow7, 2))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With

Set myRange = Range(Cells(LastRow + 5, 3), Cells(LastRow7, KolumnaDzis - 1))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, CELL.Column), Cells(LastRow, CELL.Column)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL

Set myRange = Range(Cells(LastRow + 5, KolumnaDzis + 4), Cells(LastRow7, KolumnaDzis + 4))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, KolumnaDzis + 4), Cells(LastRow, KolumnaDzis + 4)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, KolumnaDzis + 5), Cells(LastRow7, KolumnaDzis + 5))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, KolumnaDzis + 5), Cells(LastRow, KolumnaDzis + 5)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, KolumnaDzis + 6), Cells(LastRow7, KolumnaDzis + 6))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, KolumnaDzis + 6), Cells(LastRow, KolumnaDzis + 6)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, KolumnaDzis + 7), Cells(LastRow7, KolumnaDzis + 7))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, KolumnaDzis + 7), Cells(LastRow, KolumnaDzis + 7)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Range(Cells(2, 1), Cells(2, LastColumn + 10)).Copy
With Range(Cells(LastRow + 5, 1), Cells(LastRow7, LastColumn + 10))
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    .Font.ColorIndex = 0
End With
Application.CutCopyMode = False
Set myRange = Range(Cells(LastRow + 5, LastColumn + 4), Cells(LastRow7, LastColumn + 4))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 4), Cells(LastRow, LastColumn + 4)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 5), Cells(LastRow7, LastColumn + 5))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 5), Cells(LastRow, LastColumn + 5)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 6), Cells(LastRow7, LastColumn + 6))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 6), Cells(LastRow, LastColumn + 6)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 7), Cells(LastRow7, LastColumn + 7))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 7), Cells(LastRow, LastColumn + 7)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 8), Cells(LastRow7, LastColumn + 8))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 8), Cells(LastRow, LastColumn + 8)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 9), Cells(LastRow7, LastColumn + 9))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 9), Cells(LastRow, LastColumn + 9)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 10), Cells(LastRow7, LastColumn + 10))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Range(Cells(2, LastColumn + 10), Cells(LastRow, LastColumn + 10)), Range(Cells(2, 1), Cells(LastRow, 1)), "=" & Cells(CELL.Row, 1), Range(Cells(2, 2), Cells(LastRow, 2)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
Range(Rows(1), Rows(LastRow + 3)).EntireRow.Delete
Range(Columns(KolumnaDzis), Columns(KolumnaDzis + 3)).EntireColumn.Hidden = True
Range(Columns(LastColumn + 4), Columns(LastColumn + 5)).EntireColumn.Hidden = True
Columns(LastColumn + 6).Insert
Columns(LastColumn + 6).Insert
Columns(LastColumn + 6).Insert
Columns(LastColumn + 6).Insert
Cells(1, LastColumn + 6).Value = "PLAN"
Cells(1, LastColumn + 7).Value = "DISPATCHED"
Cells(1, LastColumn + 8).Value = "REMAIN"
Cells(1, LastColumn + 9).Value = "Shipppable"
Columns(LastColumn + 6).ColumnWidth = 10
Columns(LastColumn + 7).ColumnWidth = 11
Columns(LastColumn + 8).ColumnWidth = 11
Columns(LastColumn + 9).ColumnWidth = 11
Columns(1).ColumnWidth = 4
LastRow7 = Cells(Rows.Count, 1).End(xlUp).Row

Set myRange = Range(Cells(2, LastColumn + 6), Cells(LastRow7, LastColumn + 6))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.Sum(Range(Cells(CELL.Row, 3), Cells(CELL.Row, KolumnaDzis - 1)))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL

Worksheets("Shipment Status").Columns(5).Insert
Worksheets("Shipment Status").Cells(1, 5).Value = "Inch"
Set myRange = Worksheets("Shipment Status").Range(Worksheets("Shipment Status").Cells(2, 5), Worksheets("Shipment Status").Cells(LastRow5, 5))
For Each CELL In myRange
CELL.Value = Mid(CELL.Offset(0, 1).Value, 3, 2)
Next CELL
Set myRange = Range(Cells(2, LastColumn + 7), Cells(LastRow7, LastColumn + 7))
For Each CELL In myRange
   mySum = Application.WorksheetFunction.SumIfs(Worksheets("Shipment Status").Range(Worksheets("Shipment Status").Cells(2, 24), Worksheets("Shipment Status").Cells(LastRow5, 24)), Worksheets("Shipment Status").Range(Worksheets("Shipment Status").Cells(2, 5), Worksheets("Shipment Status").Cells(LastRow5, 5)), "=" & Cells(CELL.Row, 1), Worksheets("Shipment Status").Range(Worksheets("Shipment Status").Cells(2, 12), Worksheets("Shipment Status").Cells(LastRow5, 12)), "=" & Cells(CELL.Row, 2))
     If mySum <> 0 Then CELL.Value = mySum
Next CELL
 
For Each CELL In Range(Cells(2, LastColumn + 8), Cells(LastRow7, LastColumn + 8))
    CELL.Value = (CELL.Offset(0, -1).Value - CELL.Offset(0, -2).Value)
Next CELL

With Range(Cells(2, LastColumn + 8), Cells(LastRow7, LastColumn + 8))
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
        With .FormatConditions(1)
                .Interior.Color = RGB(234, 234, 234)
                .Font.Bold = True
                .Font.Italic = False
                .Font.ColorIndex = 3
        End With
    .FormatConditions(2).Font.Bold = True
End With
For Each CELL In Range(Cells(2, LastColumn + 9), Cells(LastRow7, LastColumn + 9))
    CELL.Value = CELL.Offset(0, -5).Value
Next CELL
 
Worksheets(Worksheets.Count).Name = "Shipment Summary per Inch"
Worksheets(Worksheets.Count).Activate
Worksheets(Worksheets.Count).Cells(1, 1).Select
Worksheets(Worksheets.Count - 1).Name = "Shipment Summary per Buyer"
Worksheets(Worksheets.Count - 1).Activate
Worksheets(Worksheets.Count - 1).Cells(1, 1).Select
End Sub
Private Sub Kolor_Lampki_1()
Set myRange = Range(Cells(LastRow + 5, LastColumn + 15), Cells(LastRow2 - 1, LastColumn + 15))
For Each CELL In myRange
i = CELL.Row
        If CELL.Value < 0 Then
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) > 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            Cells(i, KolumnaDzis).Interior.ColorIndex = 3
                                            Cells(i, KolumnaDzis).Font.ColorIndex = 2
                                            Cells(i, KolumnaDzis).Font.Bold = True
                                            Cells(i, 1).Interior.ColorIndex = 3
                                            Cells(i, 1).Font.ColorIndex = 2
                                            Cells(i, 1).Font.Bold = True
                        End If
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) < 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            Cells(i, KolumnaDzis).Interior.ColorIndex = 5
                                            Cells(i, KolumnaDzis).Font.ColorIndex = 2
                                            Cells(i, KolumnaDzis).Font.Bold = True
                                            Cells(i, 1).Interior.ColorIndex = 5
                                            Cells(i, 1).Font.ColorIndex = 2
                                            Cells(i, 1).Font.Bold = True
                                           
                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value < 0 Then
                                                    If Cells(i, KolumnaDzis + 5).Value > 0 Then
                                                            Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 5).Font.Bold = True
                                                    End If
                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value < 0 Then
                                                            If Cells(i, KolumnaDzis + 6).Value > 0 Then
                                                                    Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                            End If
                                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value < 0 Then
                                                                    If Cells(i, KolumnaDzis + 7).Value > 0 Then
                                                                            Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                                    End If
                                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value + Cells(i, KolumnaDzis + 8).Value < 0 Then
                                                                            If Cells(i, KolumnaDzis + 8).Value > 0 Then
                                                                                    Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                                    Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                                    Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                            End If
                                                                    Else
                                                                            Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                    End If
                                                            Else
                                                                    Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                            End If
                                                           
                                                    Else
                                                            Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                    End If
                                            Else
                                                    Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 5).Font.Bold = True
                                            End If
                        End If
                        If WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) = 0 And WorksheetFunction.Sum(Cells(i, KolumnaDzis + 5), Cells(i, KolumnaDzis + 6), Cells(i, KolumnaDzis + 7), Cells(i, KolumnaDzis + 8)) = 0 Then
                                If Cells(i, KolumnaDzis).Value > 0 Then
                                            Cells(i, KolumnaDzis).Interior.ColorIndex = 6
                                            Cells(i, KolumnaDzis).Font.ColorIndex = 1
                                            Cells(i, KolumnaDzis).Font.Bold = True
                                            Cells(i, 1).Interior.ColorIndex = 6
                                            Cells(i, 1).Font.ColorIndex = 1
                                            Cells(i, 1).Font.Bold = True
                                End If
                        End If
        End If
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 16), Cells(LastRow2 - 1, LastColumn + 16))
For Each CELL In myRange
i = CELL.Row
        If CELL.Value < 0 Then
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) > 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            If Cells(i, KolumnaDzis + 1).Value > 0 Then
                                                        Cells(i, KolumnaDzis + 1).Interior.ColorIndex = 3
                                                        Cells(i, KolumnaDzis + 1).Font.ColorIndex = 2
                                                        Cells(i, KolumnaDzis + 1).Font.Bold = True
                                                        Cells(i, 1).Interior.ColorIndex = 3
                                                        Cells(i, 1).Font.ColorIndex = 2
                                                        Cells(i, 1).Font.Bold = True
                                            End If
                        End If
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) < 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            If Cells(i, KolumnaDzis + 1).Value > 0 Then
                                                    Cells(i, KolumnaDzis + 1).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 1).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 1).Font.Bold = True
                                                    Cells(i, 1).Interior.ColorIndex = 5
                                                    Cells(i, 1).Font.ColorIndex = 2
                                                    Cells(i, 1).Font.Bold = True
                                            End If
                                           
                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value < 0 Then
                                                    If Cells(i, KolumnaDzis + 5).Value > 0 Then
                                                            Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 5).Font.Bold = True
                                                    End If
                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value < 0 Then
                                                            If Cells(i, KolumnaDzis + 6).Value > 0 Then
                                                                    Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                            End If
                                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value < 0 Then
                                                                    If Cells(i, KolumnaDzis + 7).Value > 0 Then
                                                                            Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                                    End If
                                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value + Cells(i, KolumnaDzis + 8).Value < 0 Then
                                                                            If Cells(i, KolumnaDzis + 8).Value > 0 Then
                                                                                    Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                                    Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                                    Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                            End If
                                                                    Else
                                                                            Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                    End If
                                                            Else
                                                                    Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                            End If
                                                           
                                                    Else
                                                            Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                    End If
                                            Else
                                                    Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 5).Font.Bold = True
                                            End If
                        End If
                        If WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) = 0 And WorksheetFunction.Sum(Cells(i, KolumnaDzis + 5), Cells(i, KolumnaDzis + 6), Cells(i, KolumnaDzis + 7), Cells(i, KolumnaDzis + 8)) = 0 Then
                                If Cells(i, KolumnaDzis + 1).Value > 0 Then
                                            Cells(i, KolumnaDzis + 1).Interior.ColorIndex = 6
                                            Cells(i, KolumnaDzis + 1).Font.ColorIndex = 1
                                            Cells(i, KolumnaDzis + 1).Font.Bold = True
                                            Cells(i, 1).Interior.ColorIndex = 6
                                            Cells(i, 1).Font.ColorIndex = 1
                                            Cells(i, 1).Font.Bold = True
                                End If
                        End If
        End If
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 17), Cells(LastRow2 - 1, LastColumn + 17))
For Each CELL In myRange
i = CELL.Row
        If CELL.Value < 0 Then
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) > 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            If Cells(i, KolumnaDzis + 2).Value > 0 Then
                                                        Cells(i, KolumnaDzis + 2).Interior.ColorIndex = 3
                                                        Cells(i, KolumnaDzis + 2).Font.ColorIndex = 2
                                                        Cells(i, KolumnaDzis + 2).Font.Bold = True
                                                        Cells(i, 1).Interior.ColorIndex = 3
                                                        Cells(i, 1).Font.ColorIndex = 2
                                                        Cells(i, 1).Font.Bold = True
                                            End If
                        End If
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) < 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            If Cells(i, KolumnaDzis + 2).Value > 0 Then
                                                    Cells(i, KolumnaDzis + 2).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 2).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 2).Font.Bold = True
                                                    Cells(i, 1).Interior.ColorIndex = 5
                                                    Cells(i, 1).Font.ColorIndex = 2
                                                    Cells(i, 1).Font.Bold = True
                                            End If
                                           
                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value < 0 Then
                                                    If Cells(i, KolumnaDzis + 5).Value > 0 Then
                                                            Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 5).Font.Bold = True
                                                    End If
                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value < 0 Then
                                                            If Cells(i, KolumnaDzis + 6).Value > 0 Then
                                                                    Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                            End If
                                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value < 0 Then
                                                                    If Cells(i, KolumnaDzis + 7).Value > 0 Then
                                                                            Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                                    End If
                                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value + Cells(i, KolumnaDzis + 8).Value < 0 Then
                                                                            If Cells(i, KolumnaDzis + 8).Value > 0 Then
                                                                                    Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                                    Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                                    Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                            End If
                                                                    Else
                                                                            Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                    End If
                                                            Else
                                                                    Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                            End If
                                                           
                                                    Else
                                                            Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                    End If
                                            Else
                                                    Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 5).Font.Bold = True
                                            End If
                        End If
                        If WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) = 0 And WorksheetFunction.Sum(Cells(i, KolumnaDzis + 5), Cells(i, KolumnaDzis + 6), Cells(i, KolumnaDzis + 7), Cells(i, KolumnaDzis + 8)) = 0 Then
                                If Cells(i, KolumnaDzis + 2).Value > 0 Then
                                            Cells(i, KolumnaDzis + 2).Interior.ColorIndex = 6
                                            Cells(i, KolumnaDzis + 2).Font.ColorIndex = 1
                                            Cells(i, KolumnaDzis + 2).Font.Bold = True
                                            Cells(i, 1).Interior.ColorIndex = 6
                                            Cells(i, 1).Font.ColorIndex = 1
                                            Cells(i, 1).Font.Bold = True
                                End If
                        End If
        End If
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 18), Cells(LastRow2 - 1, LastColumn + 18))
For Each CELL In myRange
i = CELL.Row
        If CELL.Value < 0 Then
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) > 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            If Cells(i, KolumnaDzis + 3).Value > 0 Then
                                                        Cells(i, KolumnaDzis + 3).Interior.ColorIndex = 3
                                                        Cells(i, KolumnaDzis + 3).Font.ColorIndex = 2
                                                        Cells(i, KolumnaDzis + 3).Font.Bold = True
                                                        Cells(i, 1).Interior.ColorIndex = 3
                                                        Cells(i, 1).Font.ColorIndex = 2
                                                        Cells(i, 1).Font.Bold = True
                                            End If
                        End If
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) < 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            If Cells(i, KolumnaDzis + 3).Value > 0 Then
                                                    Cells(i, KolumnaDzis + 3).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 3).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 3).Font.Bold = True
                                                    Cells(i, 1).Interior.ColorIndex = 5
                                                    Cells(i, 1).Font.ColorIndex = 2
                                                    Cells(i, 1).Font.Bold = True
                                            End If
                                           
                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value < 0 Then
                                                    If Cells(i, KolumnaDzis + 5).Value > 0 Then
                                                            Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 5).Font.Bold = True
                                                    End If
                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value < 0 Then
                                                            If Cells(i, KolumnaDzis + 6).Value > 0 Then
                                                                    Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                            End If
                                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value < 0 Then
                                                                    If Cells(i, KolumnaDzis + 7).Value > 0 Then
                                                                            Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                                    End If
                                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value + Cells(i, KolumnaDzis + 8).Value < 0 Then
                                                                            If Cells(i, KolumnaDzis + 8).Value > 0 Then
                                                                                    Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                                    Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                                    Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                            End If
                                                                    Else
                                                                            Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                    End If
                                                            Else
                                                                    Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                            End If
                                                           
                                                    Else
                                                            Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                    End If
                                            Else
                                                    Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 5).Font.Bold = True
                                            End If
                        End If
                        If WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) = 0 And WorksheetFunction.Sum(Cells(i, KolumnaDzis + 5), Cells(i, KolumnaDzis + 6), Cells(i, KolumnaDzis + 7), Cells(i, KolumnaDzis + 8)) = 0 Then
                                If Cells(i, KolumnaDzis + 3).Value > 0 Then
                                            Cells(i, KolumnaDzis + 3).Interior.ColorIndex = 6
                                            Cells(i, KolumnaDzis + 3).Font.ColorIndex = 1
                                            Cells(i, KolumnaDzis + 3).Font.Bold = True
                                            Cells(i, 1).Interior.ColorIndex = 6
                                            Cells(i, 1).Font.ColorIndex = 1
                                            Cells(i, 1).Font.Bold = True
                                End If
                        End If
        End If
Next CELL
Set myRange = Range(Cells(LastRow + 5, LastColumn + 19), Cells(LastRow2 - 1, LastColumn + 19))
For Each CELL In myRange
i = CELL.Row
        If CELL.Value < 0 Then
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) > 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            If Cells(i, KolumnaDzis + 4).Value > 0 Then
                                                        Cells(i, KolumnaDzis + 4).Interior.ColorIndex = 3
                                                        Cells(i, KolumnaDzis + 4).Font.ColorIndex = 2
                                                        Cells(i, KolumnaDzis + 4).Font.Bold = True
                                                        Cells(i, 1).Interior.ColorIndex = 3
                                                        Cells(i, 1).Font.ColorIndex = 2
                                                        Cells(i, 1).Font.Bold = True
                                            End If
                        End If
                        If CELL.Value + WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) < 0 Then
                                            For c = 7 To 11
                                                If Cells(i, LastColumn + c).Value > 0 Then
                                                        Cells(i, LastColumn + c).Interior.ColorIndex = 3
                                                        Cells(i, LastColumn + c).Font.ColorIndex = 2
                                                        Cells(i, LastColumn + c).Font.Bold = True
                                                End If
                                            Next c
                                            If Cells(i, KolumnaDzis + 4).Value > 0 Then
                                                    Cells(i, KolumnaDzis + 4).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 4).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 4).Font.Bold = True
                                                    Cells(i, 1).Interior.ColorIndex = 5
                                                    Cells(i, 1).Font.ColorIndex = 2
                                                    Cells(i, 1).Font.Bold = True
                                            End If
                                           
                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value < 0 Then
                                                    If Cells(i, KolumnaDzis + 5).Value > 0 Then
                                                            Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 5).Font.Bold = True
                                                    End If
                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value < 0 Then
                                                            If Cells(i, KolumnaDzis + 6).Value > 0 Then
                                                                    Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                            End If
                                                            If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value < 0 Then
                                                                    If Cells(i, KolumnaDzis + 7).Value > 0 Then
                                                                            Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                                    End If
                                                                    If CELL.Value + Cells(i, KolumnaDzis + 5).Value + Cells(i, KolumnaDzis + 6).Value + Cells(i, KolumnaDzis + 7).Value + Cells(i, KolumnaDzis + 8).Value < 0 Then
                                                                            If Cells(i, KolumnaDzis + 8).Value > 0 Then
                                                                                    Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                                    Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                                    Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                            End If
                                                                    Else
                                                                            Cells(i, KolumnaDzis + 8).Interior.ColorIndex = 5
                                                                            Cells(i, KolumnaDzis + 8).Font.ColorIndex = 2
                                                                            Cells(i, KolumnaDzis + 8).Font.Bold = True
                                                                    End If
                                                            Else
                                                                    Cells(i, KolumnaDzis + 7).Interior.ColorIndex = 5
                                                                    Cells(i, KolumnaDzis + 7).Font.ColorIndex = 2
                                                                    Cells(i, KolumnaDzis + 7).Font.Bold = True
                                                            End If
                                                           
                                                    Else
                                                            Cells(i, KolumnaDzis + 6).Interior.ColorIndex = 5
                                                            Cells(i, KolumnaDzis + 6).Font.ColorIndex = 2
                                                            Cells(i, KolumnaDzis + 6).Font.Bold = True
                                                    End If
                                            Else
                                                    Cells(i, KolumnaDzis + 5).Interior.ColorIndex = 5
                                                    Cells(i, KolumnaDzis + 5).Font.ColorIndex = 2
                                                    Cells(i, KolumnaDzis + 5).Font.Bold = True
                                            End If
                        End If
                        If WorksheetFunction.Sum(Cells(i, LastColumn + 7), Cells(i, LastColumn + 8), Cells(i, LastColumn + 9), Cells(i, LastColumn + 10), Cells(i, LastColumn + 11)) = 0 And WorksheetFunction.Sum(Cells(i, KolumnaDzis + 5), Cells(i, KolumnaDzis + 6), Cells(i, KolumnaDzis + 7), Cells(i, KolumnaDzis + 8)) = 0 Then
                                If Cells(i, KolumnaDzis + 4).Value > 0 Then
                                            Cells(i, KolumnaDzis + 4).Interior.ColorIndex = 6
                                            Cells(i, KolumnaDzis + 4).Font.ColorIndex = 1
                                            Cells(i, KolumnaDzis + 4).Font.Bold = True
                                            Cells(i, 1).Interior.ColorIndex = 6
                                            Cells(i, 1).Font.ColorIndex = 1
                                            Cells(i, 1).Font.Bold = True
                                End If
                        End If
        End If
Next CELL
End Sub

'''Private Sub Kolor_Lampki_2()
'''
'''Set myRange = Range(Cells(LastRow + 5, 1), Cells(LastRow2 - 1, 1))
'''For Each CELL In myRange
'''
'''If CELL.Interior.ColorIndex = 6 Then
'''With Worksheets("Sheet1").Range(Cells(2, 1), Cells(LastRow, 1))
'''    Set myModel = .Find(CELL.Value, LookIn:=xlValues)
'''    If Not myModel Is Nothing Then
'''        firstAddress = myModel.Address
'''        Do
'''            myModel.Interior.ColorIndex = 6
'''            Set myRange2 = Range(Cells(myModel.Row, KolumnaDzis), Cells(myModel.Row, KolumnaDzis + 4))
'''            For Each CELL2 In myRange2
'''            If CELL2 > 0 Then
'''                CELL2.Interior.ColorIndex = 6
'''            End If
'''            Next CELL2
'''            Set myModel = .FindNext(myModel)
'''        Loop While Not myModel Is Nothing And myModel.Address <> firstAddress
'''    End If
'''End With
'''End If
'''Next CELL
'''
'''Set myRange = Range(Cells(LastRow + 5, 1), Cells(LastRow2 - 1, 1))
'''For Each CELL In myRange
'''
'''If CELL.Interior.ColorIndex = 3 Then
'''    With Worksheets("Sheet1").Range(Cells(2, 1), Cells(LastRow, 1))
'''        Set myModel = .Find(CELL.Value, LookIn:=xlValues)
'''        If Not myModel Is Nothing Then
'''            firstAddress = myModel.Address
'''            Row1 = myModel.Row
'''            Row2 = 0
'''            Do
'''                myModel.Interior.ColorIndex = 3
'''                myModel.Font.ColorIndex = 2
'''                myModel.Font.Bold = True
'''                With Range(Cells(myModel.Row, LastColumn + 7), Cells(myModel.Row, LastColumn + 11))
'''                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
'''                    With .FormatConditions(1)
'''                        .Interior.ColorIndex = 3
'''                        .Font.ColorIndex = 2
'''                        .Font.Bold = True
'''                    End With
'''                End With
'''                Set myModel = .FindNext(myModel)
'''                If Row1 < myModel.Row Then
'''                    Row2 = myModel.Row
'''                End If
'''            Loop While Not myModel Is Nothing And myModel.Address <> firstAddress
'''            If Row2 = 0 Then Row2 = Row1
'''        End If
'''        mySum = Cells(Row1, KolumnaDzis).Value
'''        mySum3 = Application.WorksheetFunction.Sum(Range(Cells(Row1, KolumnaDzis), Cells(Row2, KolumnaDzis)))
'''        mySum4 = Cells(Row1, KolumnaDzis + 1).Value
'''        mySum6 = Application.WorksheetFunction.Sum(Range(Cells(Row1, KolumnaDzis), Cells(Row2, KolumnaDzis + 1)))
'''        mySum7 = Cells(Row1, KolumnaDzis + 2).Value
'''        mySum9 = Application.WorksheetFunction.Sum(Range(Cells(Row1, KolumnaDzis), Cells(Row2, KolumnaDzis + 2)))
'''        mySum10 = Cells(Row1, KolumnaDzis + 3).Value
'''        mySum12 = Application.WorksheetFunction.Sum(Range(Cells(Row1, KolumnaDzis), Cells(Row2, KolumnaDzis + 3)))
'''        mySum13 = Cells(Row1, KolumnaDzis + 4).Value
'''        mySum15 = Application.WorksheetFunction.Sum(Range(Cells(Row1, KolumnaDzis), Cells(Row2, KolumnaDzis + 4)))
'''
'''        If Row1 = Row2 Then
'''                If Cells(CELL.Row, LastColumn + 15) < 0 Then
'''                    If mySum > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis) > 0 Then
'''                                Cells(Row1, KolumnaDzis).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 16) < 0 Then
'''                    If mySum6 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 1) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 1).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 1).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 1).Font.Bold = True
'''
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 17) < 0 Then
'''                    If mySum9 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 2) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 2).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 2).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 2).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 18) < 0 Then
'''                    If mySum12 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 3) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 3).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 3).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 3).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 19) < 0 Then
'''                    If mySum15 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 4) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 4).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 4).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 4).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''
'''        End If
'''        If Row2 - Row1 = 1 Then
'''                If Cells(CELL.Row, LastColumn + 15) < 0 Then
'''                    If mySum > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis) > 0 Then
'''                                Cells(Row1, KolumnaDzis).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum3 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis) > 0 Then
'''                                Cells(Row2, KolumnaDzis).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 16) < 0 Then
'''                    If (mySum3 + mySum4) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 1) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 1).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 1).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 1).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum6 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis + 1) > 0 Then
'''                                Cells(Row2, KolumnaDzis + 1).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis + 1).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis + 1).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 17) < 0 Then
'''                    If (mySum6 + mySum7) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 2) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 2).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 2).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 2).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum9 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis + 2) > 0 Then
'''                                Cells(Row2, KolumnaDzis + 2).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis + 2).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis + 2).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 18) < 0 Then
'''                    If (mySum9 + mySum10) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 3) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 3).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 3).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 3).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum12 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis + 3) > 0 Then
'''                                Cells(Row2, KolumnaDzis + 3).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis + 3).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis + 3).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 19) < 0 Then
'''                    If (mySum12 + mySum13) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 4) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 4).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 4).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 4).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum15 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis + 4) > 0 Then
'''                                Cells(Row2, KolumnaDzis + 4).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis + 4).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis + 4).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(Row1, KolumnaDzis).Interior.ColorIndex <> 3 And Cells(Row1, KolumnaDzis + 1).Interior.ColorIndex <> 3 And Cells(Row1, KolumnaDzis + 2).Interior.ColorIndex <> 3 And Cells(Row1, KolumnaDzis + 3).Interior.ColorIndex <> 3 And Cells(Row1, KolumnaDzis + 4).Interior.ColorIndex <> 3 Then
'''                    Cells(Row1, 1).Interior.ColorIndex = 2
'''                    Cells(Row1, 1).Font.ColorIndex = 1
'''                    Range(Cells(Row1, LastColumn + 7), Cells(Row1, LastColumn + 11)).FormatConditions.Delete
'''                End If
'''                If Cells(Row2, KolumnaDzis).Interior.ColorIndex <> 3 And Cells(Row2, KolumnaDzis + 1).Interior.ColorIndex <> 3 And Cells(Row2, KolumnaDzis + 2).Interior.ColorIndex <> 3 And Cells(Row2, KolumnaDzis + 3).Interior.ColorIndex <> 3 And Cells(Row2, KolumnaDzis + 4).Interior.ColorIndex <> 3 Then
'''                    Cells(Row2, 1).Interior.ColorIndex = 2
'''                    Cells(Row2, 1).Font.ColorIndex = 1
'''                    Range(Cells(Row2, LastColumn + 7), Cells(Row2, LastColumn + 11)).FormatConditions.Delete
'''                End If
'''
'''        End If
'''        If Row2 - Row1 = 2 Then
'''                mySum2 = Cells(Row1 + 1, KolumnaDzis).Value
'''                mySum5 = Cells(Row1 + 1, KolumnaDzis + 1).Value
'''                mySum8 = Cells(Row1 + 1, KolumnaDzis + 2).Value
'''                mySum11 = Cells(Row1 + 1, KolumnaDzis + 3).Value
'''                mySum14 = Cells(Row1 + 1, KolumnaDzis + 4).Value
'''
'''                If Cells(CELL.Row, LastColumn + 15) < 0 Then
'''                    If mySum > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis) > 0 Then
'''                                Cells(Row1, KolumnaDzis).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis).Font.Bold = True
'''                            End If
'''                    End If
'''                    If (mySum + mySum2) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1 + 1, KolumnaDzis) > 0 Then
'''                                Cells(Row1 + 1, KolumnaDzis).Interior.ColorIndex = 3
'''                                Cells(Row1 + 1, KolumnaDzis).Font.ColorIndex = 2
'''                                Cells(Row1 + 1, KolumnaDzis).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum3 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis) > 0 Then
'''                                Cells(Row2, KolumnaDzis).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 16) < 0 Then
'''                    If (mySum3 + mySum4) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 1) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 1).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 1).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 1).Font.Bold = True
'''                            End If
'''                    End If
'''                    If (mySum3 + mySum4 + mySum5) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1 + 1, KolumnaDzis + 1) > 0 Then
'''                                Cells(Row1 + 1, KolumnaDzis + 1).Interior.ColorIndex = 3
'''                                Cells(Row1 + 1, KolumnaDzis + 1).Font.ColorIndex = 2
'''                                Cells(Row1 + 1, KolumnaDzis + 1).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum6 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis + 1) > 0 Then
'''                                Cells(Row2, KolumnaDzis + 1).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis + 1).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis + 1).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 17) < 0 Then
'''                    If (mySum6 + mySum7) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 2) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 2).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 2).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 2).Font.Bold = True
'''                            End If
'''                    End If
'''                    If (mySum6 + mySum7 + mySum8) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1 + 1, KolumnaDzis + 2) > 0 Then
'''                                Cells(Row1 + 1, KolumnaDzis + 2).Interior.ColorIndex = 3
'''                                Cells(Row1 + 1, KolumnaDzis + 2).Font.ColorIndex = 2
'''                                Cells(Row1 + 1, KolumnaDzis + 2).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum9 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis + 2) > 0 Then
'''                                Cells(Row2, KolumnaDzis + 2).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis + 2).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis + 2).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 18) < 0 Then
'''                    If (mySum9 + mySum10) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 3) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 3).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 3).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 3).Font.Bold = True
'''                            End If
'''                    End If
'''                    If (mySum9 + mySum10 + mySum11) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1 + 1, KolumnaDzis + 3) > 0 Then
'''                                Cells(Row1 + 1, KolumnaDzis + 3).Interior.ColorIndex = 3
'''                                Cells(Row1 + 1, KolumnaDzis + 3).Font.ColorIndex = 2
'''                                Cells(Row1 + 1, KolumnaDzis + 3).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum12 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis + 3) > 0 Then
'''                                Cells(Row2, KolumnaDzis + 3).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis + 3).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis + 3).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(CELL.Row, LastColumn + 19) < 0 Then
'''                    If (mySum12 + mySum13) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1, KolumnaDzis + 4) > 0 Then
'''                                Cells(Row1, KolumnaDzis + 4).Interior.ColorIndex = 3
'''                                Cells(Row1, KolumnaDzis + 4).Font.ColorIndex = 2
'''                                Cells(Row1, KolumnaDzis + 4).Font.Bold = True
'''                            End If
'''                    End If
'''                    If (mySum12 + mySum13 + mySum14) > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row1 + 1, KolumnaDzis + 4) > 0 Then
'''                                Cells(Row1 + 1, KolumnaDzis + 4).Interior.ColorIndex = 3
'''                                Cells(Row1 + 1, KolumnaDzis + 4).Font.ColorIndex = 2
'''                                Cells(Row1 + 1, KolumnaDzis + 4).Font.Bold = True
'''                            End If
'''                    End If
'''                    If mySum15 > (Cells(Row1, LastColumn + 5) + Cells(Row1, LastColumn + 6)) Then
'''                            If Cells(Row2, KolumnaDzis + 4) > 0 Then
'''                                Cells(Row2, KolumnaDzis + 4).Interior.ColorIndex = 3
'''                                Cells(Row2, KolumnaDzis + 4).Font.ColorIndex = 2
'''                                Cells(Row2, KolumnaDzis + 4).Font.Bold = True
'''                            End If
'''                    End If
'''                End If
'''                If Cells(Row1, KolumnaDzis).Interior.ColorIndex <> 3 And Cells(Row1, KolumnaDzis + 1).Interior.ColorIndex <> 3 And Cells(Row1, KolumnaDzis + 2).Interior.ColorIndex <> 3 And Cells(Row1, KolumnaDzis + 3).Interior.ColorIndex <> 3 And Cells(Row1, KolumnaDzis + 4).Interior.ColorIndex <> 3 Then
'''                    Cells(Row1, 1).Interior.ColorIndex = 2
'''                    Cells(Row1, 1).Font.ColorIndex = 1
'''                    Range(Cells(Row1, LastColumn + 7), Cells(Row1, LastColumn + 11)).FormatConditions.Delete
'''                End If
'''                If Cells(Row1 + 1, KolumnaDzis).Interior.ColorIndex <> 3 And Cells(Row1 + 1, KolumnaDzis + 1).Interior.ColorIndex <> 3 And Cells(Row1 + 1, KolumnaDzis + 2).Interior.ColorIndex <> 3 And Cells(Row1 + 1, KolumnaDzis + 3).Interior.ColorIndex <> 3 And Cells(Row1 + 1, KolumnaDzis + 4).Interior.ColorIndex <> 3 Then
'''                    Cells(Row1 + 1, 1).Interior.ColorIndex = 2
'''                    Cells(Row1 + 1, 1).Font.ColorIndex = 1
'''                    Range(Cells(Row1 + 1, LastColumn + 7), Cells(Row1 + 1, LastColumn + 11)).FormatConditions.Delete
'''                End If
'''                If Cells(Row2, KolumnaDzis).Interior.ColorIndex <> 3 And Cells(Row2, KolumnaDzis + 1).Interior.ColorIndex <> 3 And Cells(Row2, KolumnaDzis + 2).Interior.ColorIndex <> 3 And Cells(Row2, KolumnaDzis + 3).Interior.ColorIndex <> 3 And Cells(Row2, KolumnaDzis + 4).Interior.ColorIndex <> 3 Then
'''                    Cells(Row2, 1).Interior.ColorIndex = 2
'''                    Cells(Row2, 1).Font.ColorIndex = 1
'''                    Range(Cells(Row2, LastColumn + 7), Cells(Row2, LastColumn + 11)).FormatConditions.Delete
'''                End If
'''
'''        End If
'''
'''    End With
'''End If
'''Next CELL
'''
'''Set myRange = Range(Cells(LastRow + 5, 1), Cells(LastRow2 - 1, 1))
'''For Each CELL In myRange
'''
'''If CELL.Interior.ColorIndex = 5 Then
'''With Worksheets("Sheet1").Range(Cells(2, 1), Cells(LastRow, 1))
'''    Set myModel = .Find(CELL.Value, LookIn:=xlValues)
'''    If Not myModel Is Nothing Then
'''        firstAddress = myModel.Address
'''        Do
'''            myModel.Interior.ColorIndex = 5
'''            myModel.Font.ColorIndex = 2
'''            myModel.Font.Bold = True
'''            With Range(Cells(myModel.Row, LastColumn + 7), Cells(myModel.Row, LastColumn + 11))
'''                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
'''                 With .FormatConditions(1)
'''                        .Interior.ColorIndex = 3
'''                        .Font.ColorIndex = 2
'''                        .Font.Bold = True
'''                 End With
'''            End With
'''
'''            Set myModel = .FindNext(myModel)
'''        Loop While Not myModel Is Nothing And myModel.Address <> firstAddress
'''    End If
'''End With
'''End If
'''Next CELL

'''End Sub


