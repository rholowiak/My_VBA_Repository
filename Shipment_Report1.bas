Attribute VB_Name = "Module1"
Sub Shipment_Report1()

 

    ' Shipment delivery and inventory status Report

    ' Written by Rafal Holowiak 2010-11-09

 

Dim CELL As Range, CELL2 As Range, CELL3 As Range, CELL5 As Range, myRange As Range, n As Range, CELL6 As Range, CELL7 As Range, CELL8 As Range, CELL9 As Range

Dim LastRow As Integer, LastRow2 As Integer, LastRow3 As Integer, LastRow4 As Integer, LastRow5 As Integer, KolumnaHoldReason As Integer

Dim LastColumn As Integer, KolumnaDzisProduction As Integer, KolumnaDzis As Integer, KolumnaHold As Integer, KolumnaModel As Integer

Dim nResult As Integer, NRofLastList As Integer, SumAi As Integer, h As Integer, i As Integer, c As Integer

Dim LastColumn5 As Integer, LastRow6 As Integer, LastRow7 As Integer, myRow5 As Integer

Dim mySum As Long, mojaZmienna1 As Long, mojaZmienna2 As Long, mojaZmienna3 As Long, mojaZmienna4 As Long, myShippingTotal As Long

Dim FDofMonth As Date, FDofMonth2 As Date

Dim nStart(2) As String, nQuit(2) As String, nAccept(2) As String, Suffix3 As String, firstAddress As String

Dim M As Integer, Y As Integer

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

Application.AddCustomList ListArray:=Array("PANASONIC", "TOSHIBA", "SONY", "LGE", "PHILIPS", "ETC-TR", _

        "VESTEL E", "ARCELIK", "LOEWE A", "METZ", "WISTRON O")

NRofLastList = Application.CustomListCount + 1

Worksheets("Sheet1").Cells(1, 1).CurrentRegion.Sort Key1:=Range("B2"), Order1:=xlAscending, Key2:=Range("A2") _

        , Order2:=xlAscending, Header:=xlYes, OrderCustom:=NRofLastList, MatchCase:=True _

        , Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:= _

        xlSortNormal

 

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

For Each CELL2 In Worksheets("PalletList").Range(Worksheets("PalletList").Cells(2, 255), _

                                                                Worksheets("PalletList").Cells(LastRow3, 255))

   

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

For Each CELL2 In Worksheets("PalletList").Range(Worksheets("PalletList").Cells(2, 256), _

                                                                Worksheets("PalletList").Cells(LastRow3, 256))

 

        CELL2.Value = CELL2.Offset(, -254).Value & CELL2.Offset(, -244).Value & CELL2.Offset(, -1).Value

 

Next CELL2

For Each CELL2 In Worksheets("PalletList").Range(Worksheets("PalletList").Cells(2, 254), _

                                                                Worksheets("PalletList").Cells(LastRow3, 254))

 

        CELL2.Value = CELL2.Offset(, -252).Value & CELL2.Offset(, 1).Value

 

Next CELL2

For Each CELL In Range(Cells(2, LastColumn + 1), Cells(LastRow, LastColumn + 1))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, 256 - LastColumn - 2), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

 

Next CELL

For Each CELL In Range(Cells(2, LastColumn + 2), Cells(LastRow, LastColumn + 2))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 4)), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

 

Next CELL

For Each CELL In Range(Cells(2, LastColumn + 3), Cells(LastRow, LastColumn + 3))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 6)), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

 

Next CELL

For Each CELL In Range(Cells(2, LastColumn + 4), Cells(LastRow, LastColumn + 4))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 8)), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

 

Next CELL

For Each CELL In Range(Cells(2, LastColumn + 5), Cells(LastRow, LastColumn + 5))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 10)), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

 

Next CELL

For Each CELL In Range(Cells(2, LastColumn + 6), Cells(LastRow, LastColumn + 6))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 12)), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

 

Next CELL

For Each CELL In Range(Cells(2, LastColumn + 7), Cells(LastRow, LastColumn + 7))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 256), .Cells(LastRow3, 256)), CELL.Offset(0, (256 - LastColumn - 14)), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

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

        , OrderCustom:=1, MatchCase:=True, Orientation:=xlTopToBottom, _

        DataOption1:=xlSortNormal

 

Cells(LastRow2 + 1, 1).Value = "TOTAL"

LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row

Range(Cells(1, 2), Cells(1, LastColumn)).Copy Range(Cells(LastRow + 4, 2), Cells(LastRow + 4, 2))

Cells(LastRow + 3, 1) = "Calculation by Model Name"

Cells(LastRow + 3, 2) = "All"

Cells(LastRow + 3, 3) = "All"

Range(Cells(1, 1), Cells(1, 3)).Copy

With Range(Cells(LastRow + 3, 1), Cells(LastRow + 3, 3))

    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _

        SkipBlanks:=False, Transpose:=False

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

    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _

        SkipBlanks:=False, Transpose:=False

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

 

KolumnaDzis = Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, _

    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Column

 

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

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 2), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

Next CELL

For Each CELL In Range(Cells(LastRow + 5, LastColumn + 2), Cells(LastRow2, LastColumn + 2))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 4), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

 

Next CELL

For Each CELL In Range(Cells(LastRow + 5, LastColumn + 3), Cells(LastRow2, LastColumn + 3))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 6), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

Next CELL

For Each CELL In Range(Cells(LastRow + 5, LastColumn + 4), Cells(LastRow2, LastColumn + 4))

   

    c = CELL.Column

   

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 8), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

Next CELL

For Each CELL In Range(Cells(LastRow + 5, LastColumn + 5), Cells(LastRow2, LastColumn + 5))

    c = CELL.Column

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 10), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

Next CELL

For Each CELL In Range(Cells(LastRow + 5, LastColumn + 6), Cells(LastRow2, LastColumn + 6))

    c = CELL.Column

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 12), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

    End With

Next CELL

For Each CELL In Range(Cells(LastRow + 5, LastColumn + 7), Cells(LastRow2, LastColumn + 7))

    c = CELL.Column

    With Worksheets("PalletList")

        CELL.Value = Application.WorksheetFunction. _

            SumIf(.Range(.Cells(2, 254), .Cells(LastRow3, 254)), CELL.Offset(0, 256 - LastColumn - 14), _

                .Range(.Cells(2, 3), .Cells(LastRow3, 3)))

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

                            .Range(Worksheets("Shipment Status").Cells(2, 23), Worksheets("Shipment Status") _

                                .Cells(LastRow5, 23)))

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

CELL6 = Cells(1, KolumnaDzis + 5).Value

CELL7 = Cells(1, KolumnaDzis + 6).Value

CELL8 = Cells(1, KolumnaDzis + 7).Value

CELL9 = Cells(1, KolumnaDzis + 8).Value

Cells(1, KolumnaDzis + 5).Value = "Production " & CELL6

Cells(1, KolumnaDzis + 6).Value = "Production " & CELL7

Cells(1, KolumnaDzis + 7).Value = "Production " & CELL8

Cells(1, KolumnaDzis + 8).Value = "Production " & CELL9

Cells(LastRow + 4, KolumnaDzis + 5).Value = "Production " & CELL6

Cells(LastRow + 4, KolumnaDzis + 6).Value = "Production " & CELL7

Cells(LastRow + 4, KolumnaDzis + 7).Value = "Production " & CELL8

Cells(LastRow + 4, KolumnaDzis + 8).Value = "Production " & CELL9

Range(Cells(1, KolumnaDzis + 5), Cells(1, KolumnaDzis + 8)).EntireColumn.ColumnWidth = 8.71

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

Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, KolumnaDzis + 5), Worksheets("Sheet1").Cells(LastRow - 1, KolumnaDzis + 5))

For Each CELL In myRange

   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction), Worksheets(5).Cells(300, KolumnaDzisProduction)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1), Worksheets(5).Range(Worksheets(5).Cells(7, 10), Worksheets(5).Cells(300, 10)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 2))

 

     If mySum <> 0 Then CELL.Value = mySum

 

Next CELL

Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, KolumnaDzis + 6), Worksheets("Sheet1").Cells(LastRow - 1, KolumnaDzis + 6))

For Each CELL In myRange

   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction + 1), Worksheets(5).Cells(300, KolumnaDzisProduction + 1)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1), Worksheets(5).Range(Worksheets(5).Cells(7, 10), Worksheets(5).Cells(300, 10)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 2))

 

     If mySum <> 0 Then CELL.Value = mySum

 

Next CELL

Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, KolumnaDzis + 7), Worksheets("Sheet1").Cells(LastRow - 1, KolumnaDzis + 7))

For Each CELL In myRange

   mySum = Application.WorksheetFunction.SumIfs(Worksheets(5).Range(Worksheets(5).Cells(7, KolumnaDzisProduction + 2), Worksheets(5).Cells(300, KolumnaDzisProduction + 2)), Worksheets(5).Range(Worksheets(5).Cells(7, 9), Worksheets(5).Cells(300, 9)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 1), Worksheets(5).Range(Worksheets(5).Cells(7, 10), Worksheets(5).Cells(300, 10)), "=" & Worksheets("Sheet1").Cells(CELL.Row, 2))

 

     If mySum <> 0 Then CELL.Value = mySum

 

Next CELL

Set myRange = Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, KolumnaDzis + 8), Worksheets("Sheet1").Cells(LastRow - 1, KolumnaDzis + 8))

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

Application.Calculation = xlCalculationAutomatic

Application.ScreenUpdating = True

MsgBox "CREATION COMPLETED", vbInformation, "Program Status"

End Sub

 

Sub KopiaShipment_Status()

 

Dim wkbName As String

 

wkbName = Application.ActiveWorkbook.Name

 

Application.ActiveSheet.Copy

 

Application.ActiveWorkbook.SaveAs Filename:="C:\Documents and Settings\LGD\Desktop\" & Left(wkbName, 26) & " Time " & Hour(Time) & "_" & Minute(Time) & ".xls", FileFormat:=-4143

 

End Sub

