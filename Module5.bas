Attribute VB_Name = "Module5"
Option Explicit

 

Private LastRow1 As Long, LastRow2 As Long, LastRow3 As Long, FirstRow1 As Long, FirstRow2 As Long

Private i As Long, j As Long

Private NowyItem1 As Range, NowyItem2 As Range

 

 

Sub Copy_Yesterday()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''skopiowanie zamówien z dnia poprzedniego

 

    LastRow1 = Cells(Rows.Count, 1).End(xlUp).Row

    For i = LastRow1 To 2 Step -1

        Set NowyItem1 = Cells(i, 1)

            If NowyItem1.Borders(xlEdgeTop).Color = RGB(255, 0, 0) And NowyItem1.Borders(xlEdgeTop).Weight = xlThick Then

                  FirstRow1 = NowyItem1.Row

                  Exit For

            End If

    Next i

   

    ''' jesli nie wiem jaki typ ma zmienna to moge uzyc ponizszej instrukcji TypeName(zmienna)

    '''MsgBox TypeName(NowyItem1)

   

    Rows(FirstRow1 & ":" & LastRow1).Copy

    Rows(LastRow1 + 1).Insert

    LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row

    Cells(LastRow2 + 1, 1).Select

    Application.CutCopyMode = False

 

End Sub

 

Sub Nowe_Zamówienia1_w_dól()

 

Application.ScreenUpdating = False

 

    LastRow1 = Cells(Rows.Count, 11).End(xlUp).Row

    LastRow3 = Cells(Rows.Count, 15).End(xlUp).Row

    LastRow1 = Application.WorksheetFunction.Max(LastRow1, LastRow3)

   

    LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row

   

    If LastRow1 = LastRow2 Then GoTo Brak_nowych_zamówien1

   

    For i = LastRow2 To 2 Step -1

        Set NowyItem1 = Cells(i, 1)

            If NowyItem1.Borders(xlEdgeTop).Color = RGB(255, 0, 0) And NowyItem1.Borders(xlEdgeTop).Weight = xlThick Then

                  FirstRow1 = NowyItem1.Row

                  Exit For

            End If

    Next i

   

    For i = LastRow2 To FirstRow1 Step -1

        Set NowyItem1 = Cells(i, 1)

            If NowyItem1.Value = "LODTE" Then

                  NowyItem1.EntireRow.Delete

                  Exit For

            End If

    Next i

    LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row

    Range(Cells(LastRow1 + 1, 9), Cells(LastRow2, 10)).Clear

 

    Range(Cells(LastRow1 + 1, 8), Cells(LastRow2, 8)).Replace What:="2043", Replacement:=Format(Date, "yyyy"), LookAt:=xlPart, _

        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _

       ReplaceFormat:=False

   

    Range(Cells(LastRow1 + 1, 8), Cells(LastRow2, 8)).Replace What:="2020", Replacement:=Format(Date, "yyyy"), LookAt:=xlPart, _

        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _

        ReplaceFormat:=False

    Range(Cells(LastRow1 + 1, 8), Cells(LastRow2, 8)).Replace What:="2025", Replacement:=Format(Date, "yyyy"), LookAt:=xlPart, _

        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _

        ReplaceFormat:=False

    Range(Cells(LastRow1 + 1, 8), Cells(LastRow2, 8)).Replace What:="2026", Replacement:=Format(Date, "yyyy"), LookAt:=xlPart, _

        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _

        ReplaceFormat:=False

   

    

    Range(Cells(LastRow1 + 1, 8), Cells(LastRow2, 8)).TextToColumns Destination:=Range("H" & LastRow1 + 1), DataType:=xlDelimited, _

        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _

        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _

        :=Array(1, 5), TrailingMinusNumbers:=True

   

    Range(Cells(LastRow1 + 1, 1), Cells(LastRow2, 1)).TextToColumns Destination:=Range("A" & LastRow1 + 1), DataType:=xlDelimited, _

        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _

        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _

        :=Array(1, 5), TrailingMinusNumbers:=True

   

    

    Columns("R:R").NumberFormat = "@"

    For i = LastRow2 To FirstRow1 Step -1

        Set NowyItem1 = Cells(i, 18)

        NowyItem1.Value = Cells(i, 2).Value & Cells(i, 3).Value & Cells(i, 4).Value & Cells(i, 6).Value

    Next i

   

    For i = FirstRow1 + 1 To LastRow1

        Set NowyItem2 = Cells(i, 18)

        Cells(i, 17).Value = 1

        For j = NowyItem2.Row To LastRow2

            If NowyItem2.Value = Cells(j + 1, 18).Value Or NowyItem2.Value = "" Then

            Cells(i, 17).Value = ""

            End If

        Next j

   

    Next i

    '''sprawdzam puste komórki i wstawiam dziedziczona wartosc 1 lub puste

    For i = FirstRow1 + 1 To LastRow1

        Set NowyItem2 = Cells(i, 6)

        If IsEmpty(NowyItem2.Value) Then NowyItem2.Offset(, 11).Value = Cells(i - 1, 6).Offset(, 11).Value

    Next i

 

    Columns(18).Clear

   

Brak_nowych_zamówien1:

    Application.ScreenUpdating = True

End Sub

 

Sub Nowe_Zamówienia2_w_góre()

 

Application.ScreenUpdating = False

 

    LastRow1 = Cells(Rows.Count, 10).End(xlUp).Row

    LastRow3 = Cells(Rows.Count, 15).End(xlUp).Row

    LastRow1 = Application.WorksheetFunction.Max(LastRow1, LastRow3)

 

    LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row

    For i = LastRow2 To 2 Step -1

        Set NowyItem1 = Cells(i, 1)

            If NowyItem1.Borders(xlEdgeTop).Color = RGB(255, 0, 0) And NowyItem1.Borders(xlEdgeTop).Weight = xlThick Then

                  FirstRow1 = NowyItem1.Row

                  Exit For

            End If

    Next i

   

    If LastRow1 = LastRow2 Then GoTo Brak_nowych_zamówien2

   

    Columns("R:R").NumberFormat = "@"

    For i = LastRow2 To FirstRow1 Step -1

        Set NowyItem1 = Cells(i, 18)

        NowyItem1.Value = Cells(i, 2).Value & Cells(i, 3).Value & Cells(i, 4).Value & Cells(i, 6).Value

    Next i

 

    For i = FirstRow1 + 1 To LastRow2

        If Cells(i, 4).Value <> "C" And Cells(i, 4).Value <> "G" And Cells(i, 4).Value <> "" Then

            Cells(i, 17).Value = 1

        End If

    Next i

   

    

    For i = LastRow2 To FirstRow1 Step -1

        Set NowyItem2 = Cells(i, 18)

        For j = NowyItem2.Row To FirstRow1 Step -1

                If NowyItem2.Value = Cells(j - 1, 18).Value And NowyItem2.Value <> "" Then

                Cells(i, 17).Value = 1

                End If

        Next j

   

    Next i

 

    Columns(18).Clear

   

    Rows("3:3").Copy

    Rows(LastRow1 + 1 & ":" & LastRow2).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _

        SkipBlanks:=False, Transpose:=False

    Application.CutCopyMode = False

 

Brak_nowych_zamówien2:

 

    For i = FirstRow1 + 1 To LastRow2

        Set NowyItem2 = Cells(i, 8)

        If Not IsEmpty(NowyItem2.Value) Then NowyItem2.Offset(, 1).FormulaR1C1 = "=RC[-1]-RC[-8]"

    Next i

   

    ''jeszcze testuje ten kod czy dobrze zaznacza rekordy jak juz bede pewny to zamiast .Select mozna wstawic .Delete

    On Error Resume Next

    Range(Cells(FirstRow1, 17), Cells(LastRow2, 17)).SpecialCells(xlCellTypeConstants).EntireRow.Select

    On Error GoTo 0

    Application.ScreenUpdating = True

End Sub

 

 

 

Option Explicit

Private Godzina_wyslania_maila As Date

Private MyStr As String, PLMy2 As String, Itemy2 As String

Private i As Long, LastRow1 As Long, LastRow2 As Long, LastRow3 As Long, FirstRow1 As Long

 

Sub Mail_TabeleZakres()

 

    Dim rng As Range

    Dim OutApp As Object

    Dim OutMail As Object

    Dim NowyItem1 As Range

 

    'Jezeli do zmiennej typu Date chcesz recznie przypisac okreslona date,

    'musisz otoczyc ja z obu stron znakiem #,

    'a poszczególne skladniki daty oddzielic od siebie znakiem myslnika (-) lub slasha (/).

    ' przyklady data = #21-04-2010# lub data = #4/21/2010# lub czas = #23:15:20#

    ''''' stronka z funkcjami Date and Time - Date and Time Functions in VBA

    '''http://www.classanytime.com/mis333k/sjdatetime.html

    

    Godzina_wyslania_maila = #6:30:00 PM#

   

    LastRow1 = Cells(Rows.Count, 1).End(xlUp).Row

    LastRow2 = Cells(Rows.Count, 10).End(xlUp).Row

    LastRow1 = Application.WorksheetFunction.Max(LastRow1, LastRow2)

   

    For i = LastRow1 To 2 Step -1

        Set NowyItem1 = Cells(i, 1)

            If NowyItem1.Borders(xlEdgeTop).Color = RGB(255, 0, 0) And NowyItem1.Borders(xlEdgeTop).Weight = xlThick Then

                  FirstRow1 = NowyItem1.Row

                  Exit For

            End If

    Next i

 

    ''' wypelniam puste komórki data z wiersza powyzej

    ''' jezeli nie ma pustych to wlaczam obluge bledów poprzez resume next

    On Error Resume Next

    Range(Cells(FirstRow1, 1), Cells(LastRow1, 1)).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"

    On Error GoTo 0

    '''zaznaczam nazwy PLM na kolor niebieski'''

    Call Zaznacz_PLM_na_niebiesko

   

    '''tutaj blok ktory przypisze dane do zmiennych tablicowych i wygeneruje tekst maila'''''''''''''

    Call StworzTablice_i_tresc_maila

 

    Dim wkbName As String

    wkbName = Application.ActiveWorkbook.Name

    Application.Sheets.Copy

    Application.ActiveWorkbook.SaveAs Filename:="C:\Users\rholowia\Desktop\" & Left(wkbName, 3) & " - ATP " & Hour(Time) & "_" & Minute(Time) & ".xlsx" ', FileFormat:=-4143

    wkbName = Application.ActiveWorkbook.Name

       

    Set OutApp = CreateObject("Outlook.Application")

    Set OutMail = OutApp.CreateItem(0)

    Set rng = Range(Selection.Address)

                                 

    On Error Resume Next

    With OutMail

        ''' aby wykorzystac podpis juz istniejacy trzeba najpierw uruchomic .Display

        ''' a nastepnie do .HTMLBody = OutMail.HTMLBody <- dodac te instrukcje

        ''' uwaga! kolejnosc wywolania ma znaczenie

        .Display

        .To = "Daniel.Konecki@xyleminc.com; agnieszka.jakubowska@xyleminc.com; Daniel.Lysakowski@xyleminc.com"

        .cc = "Paolo.Bortolotto@Xyleminc.com; Adam.Papierok@Xyleminc.com; Filip.Ziolek@Xyleminc.com; tomasz.kopinski@xyleminc.com"

       

        '''funkcja MailSubject okreslajaca tytul maila

        .Subject = MailSubject

       

        '''funkcja StrBody okreslajaca tresc maila

        .HTMLBody = StrBody & RangetoHTML(rng) & OutMail.HTMLBody

        .Attachments.Add ActiveWorkbook.FullName

    End With

    '''tutaj poprzez Wait methode wstrzymuje macro na outlooku tak aby zadzialalo SendKeys i ctrl+k, które updatuje adresy mailowe

    '''czasami ctrl+k wlacza mi sie w Excelu zamiast w outlooku, przypuszczam ze macro za szybko dziala

    ''' niestety kod nie dziala dlatego wylaczam metode .Wait

    '''Application.Wait (Now + TimeValue("0:00:02"))

    On Error GoTo 0

    OutMail.Display

    Application.Wait (Now + TimeValue("0:00:01"))

    SendKeys "^k", True

   

    

    Set OutMail = Nothing

    Set OutApp = Nothing

    ''''''''''' zamykam plik pomocniczy bez makra zapisany na pulpicie'''''''''

    Application.ActiveWorkbook.Close

    ''''''''''' kasuje plik pomocniczy bez makra znika z pulpitu'''''''''

    Kill "C:\Users\rholowia\Desktop\" & wkbName

 

End Sub

 

Private Function RangetoHTML(rng As Range)

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

        .Columns(6).ColumnWidth = 19

        .Columns("A:P").EntireColumn.AutoFit

        .Rows("1:100").EntireRow.AutoFit

        Application.CutCopyMode = False

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

 

Private Sub StworzTablice_i_tresc_maila()

    Dim tablica1() As String

    Dim Rozmiar As Long

    MyStr = ""

    PLMy2 = ""

    Itemy2 = ""

   

    Range(Cells(FirstRow1, 10), Cells(LastRow1, 11)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("T" & FirstRow1), Unique:=True

    LastRow3 = Cells(Rows.Count, 20).End(xlUp).Row

   

    Cells(FirstRow1 + 1, 22).FormulaR1C1 = "=SUMIF(R" & FirstRow1 & "C10:R" & LastRow1 & "C10,RC[-2],R" & FirstRow1 & "C12:R" & LastRow1 & "C12)"

   

    '''jezeli jest jeden item to wyswietla error run time

    '''dlatego aby ominac ten error napisalem ponizszy warunek

    '''wykonuje AutoFill jezeli jest co najmniej 2 itemy

    If (FirstRow1 + 1) < LastRow3 Then

        Cells(FirstRow1 + 1, 22).AutoFill Destination:=Range(Cells(FirstRow1 + 1, 22), Cells(LastRow3, 22))

    End If

 

    Rozmiar = LastRow3 - FirstRow1

    ReDim tablica1(1 To Rozmiar, 1 To 3)

   

    For i = 1 To Rozmiar

    tablica1(i, 1) = Cells(FirstRow1 + i, 20).Value

    tablica1(i, 2) = Cells(FirstRow1 + i, 21).Value

    tablica1(i, 3) = Cells(FirstRow1 + i, 22).Value

    Next i

   

    For i = 1 To Rozmiar

        MyStr = Left(tablica1(i, 2), 3)

        If MyStr = "PLM" Then

            PLMy2 = PLMy2 & tablica1(i, 1) & " - " & tablica1(i, 3) & "pcs<br>"

        End If

    Next i

   

    For i = 1 To Rozmiar

        MyStr = Left(tablica1(i, 2), 3)

        If MyStr <> "PLM" And MyStr <> "" Then

            Itemy2 = Itemy2 & tablica1(i, 1) & " - " & tablica1(i, 3) & "pcs<br>"

        End If

    Next i

    Range(Columns(20), Columns(22)).EntireColumn.Clear

 

End Sub

Private Sub Zaznacz_PLM_na_niebiesko()

 

    For i = FirstRow1 To LastRow1

        MyStr = Left(Cells(i, 11).Value, 3)

        If MyStr = "PLM" Then

            With Cells(i, 11).Characters(Start:=1, Length:=3).Font

                .Bold = True

                .Underline = xlUnderlineStyleSingle

                .Color = RGB(0, 0, 255)

            End With

        End If

    Next i

 

End Sub

 

Private Function StrBody() As String

 

    '' sprawdzenie kolorów wklep w przegladarke CSS Color Names stronka www.w3schools.com

    If Time < Godzina_wyslania_maila Then

            StrBody = "<HTML><BODY style=font-size:11pt;font-family:Calibri;color:#334870>" & _

                    "<B>Daniel,</B><br>" & _

                    "prosze o sprawdzenie komponentów:<br>" & Itemy2 & _

                    "<B>Daniel L,</B><br>" & _

                    "prosze o podanie Best Delivery date dla ponizszych PLM potrzebnych na GWP:<br>" & PLMy2 & "<br>" & _

                    "</BODY></HTML>"

    Else

            StrBody = "<HTML><BODY style=font-size:11pt;font-family:Calibri;color:#334870>" & _

                    "<B>UPDATE</B>" & Format(Now, """ godz. """ & "h:mm") & _

                    "</BODY></HTML>"

               

    End If

 

End Function

 

Private Function MailSubject() As String

 

       If Time < Godzina_wyslania_maila Then

                ''' formatuje nazwe miesiaca i koncówke daty 1st, 2nd, 3rd itd uzywajac funkcji Format_english_Month_name'''

                MailSubject = "GWP - ATP " & ": date " & Format_english_Month_name(Now) & Format(Now, """ godz. """ & "h:mm")

        Else

                MailSubject = "GWP - ATP " & ": Update " & Format(Now, """ godz. """ & "h:mm")

        End If

 

End Function

 

Private Function Format_english_Month_name(Datum As Date) As String

Dim DD As String

Dim MM As String

Dim GetChoice As String

 

DD = Format(Datum, "dd")

MM = Choose(Month(Datum), "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")

GetChoice = Choose(DD, "st", "nd", "rd", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "st", "nd", "rd", "th", "th", "th", "th", "th", "th", "th", "st")

Format_english_Month_name = DD & GetChoice & " " & MM

 

End Function

 

Option Private Module

 

Sub ListEnvironVariables()

    Dim strEnviron As String

    Dim i As Long

    For i = 1 To 255

        strEnviron = Environ(i)

        '''With no data type identifier after the 0, it is implied to be of the native Integer (the default).

        '''Therefore this 0 is expressed as an Integer, which makes it a 16-bit 0.

        '''You should pass a Long 0, which will be a 32-bit 0, as the function requests a 32-bit long long number.

        ''''So, therefore you add a &:

        '''An & after a number or a variable means that it is a Long (which is 32-bits).

        '''0& is a 32-bit 0.

        If LenB(strEnviron) = 0& Then Exit For

        Debug.Print strEnviron

    Next

End Sub

 

Sub testEnviron()

Dim EnvString, Indx, Msg, PathLen    ' Declare variables.

Indx = 1    ' Initialize index to 1.

Do

    EnvString = Environ(Indx)    ' Get environment

    Debug.Print EnvString        ' variable.

    If Left(EnvString, 5) = "Path=" Then    ' Check PATH entry.

        PathLen = Len(Environ("PATH"))    ' Get length.

        Msg = "PATH entry = " & Indx & " and length = " & PathLen

        Exit Do

    Else

        Indx = Indx + 1    ' Not PATH entry,

    End If    ' so increment.

Loop Until EnvString = ""

If PathLen > 0 Then

    MsgBox Msg    ' Display message.

Else

    MsgBox "No PATH environment variable exists."

End If

 

End Sub

 

 

Private Sub Mojadata()

 

Debug.Print Time

Debug.Print Date

Debug.Print Now

Debug.Print Format(Date, "yyyy")

End Sub

Private Sub Nadrzedna()

    Dim A As Integer

    Dim B As Integer

    A = 111

    B = 222

    Debug.Print "PRZED = A: " & A, "B: " & B

    Call Podrzedna(A, B)

    Debug.Print "I PO =  A: " & CStr(A), "B: " & CStr(B)

End Sub

    

Sub Podrzedna(ByRef X As Integer, ByVal Y As Integer)

    X = 333

    Y = Y * 2

    suma = X + Y

    Debug.Print "SUMA =: " & suma

    Debug.Print "A: " & CStr(X), "B: " & CStr(Y)

End Sub

 

 

Private Sub Workbook_Mail()

    Dim OutApp As Object

    Dim OutMail As Object

    Dim SigString As String

    Dim Signature As Variant

   

    Set OutApp = CreateObject("Outlook.Application")

    Set OutMail = OutApp.CreateItem(0)

 

    SigString = Environ("appdata") & "\Microsoft\Signatures\Rafal1.htm"

 

    If Dir(SigString) <> "" Then

        Signature = GetBoiler(SigString)

    Else

        Signature = ""

    End If

    On Error Resume Next

    With OutMail

        .Display

        .To = ""

        .Subject = "Shipment Status " & Format(Now, "dd" & """th""" & " mmmm yyyy" & """ godz. """ & "h:mm")

        .HTMLBody = OutMail.HTMLBody & Signature & "<IMG src=""C:\Users\rholowia\AppData\Roaming\Microsoft\Signatures\Rafal1_files\image001.jpg"">"

 

        .Attachments.Add ActiveWorkbook.FullName

        ' You can add other files by uncommenting the following line.

        '.Attachments.Add ("C:\test.txt")

    End With

    On Error GoTo 0

 

    Set OutMail = Nothing

    Set OutApp = Nothing

End Sub

 

Function GetBoiler(ByVal sFile As String) As Variant

    Dim fso As Object

    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)

    GetBoiler = ts.ReadAll

    ts.Close

End Function

 

Private Sub test()

 

Set fs = CreateObject("Scripting.FileSystemObject")

''Set folder = fs.CreateFolder("c:\testfolder")

Set A = fs.CreateTextFile("c:\testfile2.txt", True)

A.WriteLine ("This is a test2.")

A.WriteLine ("this is a test drugi")

A.Close

 

End Sub

 

Option Explicit

Private MyStr As String, PLMy2 As String, Itemy2 As String

Private i As Long, LastRow1 As Long, LastRow2 As Long, LastRow3 As Long, FirstRow1 As Long, TempWB_LastRow1 As Long, TempWB_LastRow2 As Long

 

Sub Mail_CzesciWlosi()

 

    Dim rng As Range

    Dim OutApp As Object

    Dim OutMail As Object

    Dim NowyItem1 As Range

 

    LastRow1 = Cells(Rows.Count, 1).End(xlUp).Row

    LastRow2 = Cells(Rows.Count, 10).End(xlUp).Row

    LastRow1 = Application.WorksheetFunction.Max(LastRow1, LastRow2)

   

    For i = LastRow1 To 2 Step -1

        Set NowyItem1 = Cells(i, 1)

            If NowyItem1.Borders(xlEdgeTop).Color = RGB(255, 0, 0) And NowyItem1.Borders(xlEdgeTop).Weight = xlThick Then

                  FirstRow1 = NowyItem1.Row

                  Exit For

            End If

    Next i

 

   

    '''tutaj blok ktory przypisze dane do zmiennych tablicowych i wygeneruje tekst maila'''''''''''''

    Call StworzTablice_i_tresc_maila2

 

       

    Set OutApp = CreateObject("Outlook.Application")

    Set OutMail = OutApp.CreateItem(0)

    Set rng = Range(Selection.Address)

                                 

    On Error Resume Next

    With OutMail

        ''' aby wykorzystac podpis juz istniejacy trzeba najpierw uruchomic .Display

        ''' a nastepnie do .HTMLBody = OutMail.HTMLBody <- dodac te instrukcje

        ''' uwaga! kolejnosc wywolania ma znaczenie

        .Display

        .To = "Resales.Montecchio@Xyleminc.com"

        .cc = "Daniel.Konecki@xyleminc.com; agnieszka.jakubowska@xyleminc.com;"

       

        '''funkcja MailSubject okreslajaca tytul maila

        .Subject = MailSubject2

       

        '''funkcja StrBody okreslajaca tresc maila

        .HTMLBody = StrBody2 & RangetoHTML(rng) & OutMail.HTMLBody

 

    End With

   

    On Error GoTo 0

    OutMail.Display

    SendKeys "^k", True

   

    

    Set OutMail = Nothing

    Set OutApp = Nothing

 

End Sub

 

Private Function RangetoHTML(rng As Range)

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

        .Columns(6).ColumnWidth = 19

        .Columns("A:P").EntireColumn.AutoFit

        .Columns(1).EntireColumn.Delete

        .Columns(1).EntireColumn.Delete

        .Columns(2).EntireColumn.Delete

        .Columns(6).EntireColumn.Delete

        .Columns(10).EntireColumn.Delete

        .Rows("1:100").EntireRow.AutoFit

        Application.CutCopyMode = False

    End With

    TempWB_LastRow1 = TempWB.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row

    TempWB_LastRow2 = TempWB.Sheets(1).Cells(Rows.Count, 6).End(xlUp).Row

    TempWB_LastRow1 = Application.WorksheetFunction.Max(TempWB_LastRow1, TempWB_LastRow2)

 

    For i = 2 To TempWB_LastRow1

    If TempWB.Sheets(1).Cells(i, 11).Value <> "wlosi" And TempWB.Sheets(1).Cells(i, 3).Value = 0 Then TempWB.Sheets(1).Cells(i, 12).Value = 1

   

    Next i

 

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

 

Private Sub StworzTablice_i_tresc_maila2()

    Dim tablica1() As String

    Dim Rozmiar As Long

    MyStr = ""

    PLMy2 = ""

    Itemy2 = ""

   

    Range(Cells(FirstRow1, 10), Cells(LastRow1, 11)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("T" & FirstRow1), Unique:=True

    LastRow3 = Cells(Rows.Count, 20).End(xlUp).Row

   

    Cells(FirstRow1 + 1, 22).FormulaR1C1 = "=SUMIF(R" & FirstRow1 & "C10:R" & LastRow1 & "C10,RC[-2],R" & FirstRow1 & "C12:R" & LastRow1 & "C12)"

    Cells(FirstRow1 + 1, 23).FormulaR1C1 = "=VLOOKUP(RC[-3],R" & FirstRow1 & "C10:R" & LastRow1 & "C16,7,0)"

   

    

    '''jezeli jest jeden item to wyswietla error run time

    '''dlatego aby ominac ten error napisalem ponizszy warunek

    '''wykonuje AutoFill jezeli jest co najmniej 2 itemy

    If (FirstRow1 + 1) < LastRow3 Then

        Cells(FirstRow1 + 1, 22).AutoFill Destination:=Range(Cells(FirstRow1 + 1, 22), Cells(LastRow3, 22))

        Cells(FirstRow1 + 1, 23).AutoFill Destination:=Range(Cells(FirstRow1 + 1, 23), Cells(LastRow3, 23))

    End If

 

    Rozmiar = LastRow3 - FirstRow1

    ReDim tablica1(1 To Rozmiar, 1 To 4)

   

    For i = 1 To Rozmiar

    tablica1(i, 1) = Cells(FirstRow1 + i, 20).Value

    tablica1(i, 2) = Cells(FirstRow1 + i, 21).Value

    tablica1(i, 3) = Cells(FirstRow1 + i, 22).Value

    tablica1(i, 4) = Cells(FirstRow1 + i, 23).Value

    Next i

   

    For i = 1 To Rozmiar

        MyStr = Left(tablica1(i, 4), 5)

        If MyStr = "wlosi" Then

            Itemy2 = Itemy2 & tablica1(i, 1) & " - " & tablica1(i, 2) & " - - > " & tablica1(i, 3) & "pcs<br>"

        End If

    Next i

   

    Range(Columns(20), Columns(23)).EntireColumn.Clear

 

End Sub

Private Function StrBody2() As String

 

             StrBody2 = "<HTML><BODY style=font-size:11pt;font-family:Calibri;color:#334870>" & _

                    "<B>Hi Elisa,</B><br>" & _

                    "please give best delivery date for below items:<br>" & Itemy2

   

End Function

 

Private Function MailSubject2() As String

 

  MailSubject2 = "Delivery date for GWP components"

 

End Function

