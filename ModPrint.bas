Attribute VB_Name = "ModPrint"
Public PrintPuzNum As Integer           'Number of puzzles to print
Public PrintSud() As String             'Puzzles to print
Public PrintLev As Integer              'Yes/No print levelnames
Public Printlevnums As String           'Level numbers of the puzzles
Public PrintPage As Integer             'Number of puzzles on the page

Public Sub Init_PrintData()
    PrintPuzNum = 1
    PrintLev = 0
    PrintPage = 1
    Printlevnums = ""
End Sub

Private Sub Init_Printer()
    Printer.DrawStyle = 0               'solid lines
    Printer.ScaleMode = vbCentimeters
End Sub

Public Sub PrintPages()
    Dim StartXY(4, 1) As Double     'startpositions on page
    Dim PuzNum As Integer           'Active puzzle to print
    Dim PuzPage As Integer          'Puzzle number on page
    Dim PuzLev As Integer
    Dim CWidth As Double            'width of one cell
    Dim PageNumber As Integer
    Dim PageTxt As String
    Dim Marge As Double
    Dim Temp As Double
    Dim LeftSpacing As Integer
    Dim RightSpacing As Integer
    Dim TopSpacing As Integer
    Dim BottomSpacing As Integer
    Dim CenterPage As Double, HalfPage As Double
    Dim QrtHPage As Double, QrtVPage As Double
    Dim HorPage As Double, VerPage As Double
    Dim MaxPSize As Double
    If PrintPuzNum < 1 Then Exit Sub
    Call Init_Printer
    LeftSpacing = 2
    RightSpacing = 1
    TopSpacing = 1
    BottomSpacing = 1
    Marge = 15 / 100                 'print grid % smaller than maximum
    HorPage = Printer.ScaleWidth - LeftSpacing - RightSpacing
    VerPage = Printer.ScaleHeight - TopSpacing - BottomSpacing
    CenterPage = HorPage / 2
    QrtHPage = CenterPage / 2
    HalfPage = VerPage / 2
    QrtVPage = HalfPage / 2
    Select Case PrintPage
        Case 1
            MaxPSize = HorPage
            If HorPage > VerPage Then MaxPSize = VerPage
            CWidth = MaxPSize / GWidth
            StartXY(1, 0) = TopSpacing + HalfPage - CWidth * GWidth / 2
            StartXY(1, 1) = LeftSpacing + CenterPage - CWidth * GWidth / 2
        Case 2
            MaxPSize = HorPage
            If HorPage > HalfPage Then MaxPSize = HalfPage
            CWidth = MaxPSize / GWidth
            CWidth = CWidth - (CWidth * Marge)
            StartXY(1, 0) = TopSpacing + QrtVPage - CWidth * GWidth / 2
            StartXY(1, 1) = LeftSpacing + CenterPage - CWidth * GWidth / 2
            StartXY(2, 0) = TopSpacing + HalfPage + QrtVPage - CWidth * GWidth / 2
            StartXY(2, 1) = LeftSpacing + CenterPage - CWidth * GWidth / 2
        Case 4
            MaxPSize = CenterPage
            If CenterPage > HalfPage Then MaxPSize = HalfPage
            CWidth = MaxPSize / GWidth
            CWidth = CWidth - (CWidth * Marge)
            StartXY(1, 0) = TopSpacing + QrtVPage - CWidth * GWidth / 2
            StartXY(1, 1) = LeftSpacing + QrtHPage - CWidth * GWidth / 2
            StartXY(2, 0) = TopSpacing + QrtVPage - CWidth * GWidth / 2
            StartXY(2, 1) = LeftSpacing + CenterPage + QrtHPage - CWidth * GWidth / 2
            StartXY(3, 0) = TopSpacing + HalfPage + QrtVPage - CWidth * GWidth / 2
            StartXY(3, 1) = LeftSpacing + QrtHPage - CWidth * GWidth / 2
            StartXY(4, 0) = TopSpacing + HalfPage + QrtVPage - CWidth * GWidth / 2
            StartXY(4, 1) = LeftSpacing + CenterPage + QrtHPage - CWidth * GWidth / 2
    End Select
    PuzNum = 0
    PuzPage = 0
    PageNumber = 1
    Do
        PuzNum = PuzNum + 1
        PuzPage = PuzPage + 1
        If PuzNum > PrintPuzNum Then Exit Do
        Printer.FontBold = True     'set the fonts to Bold
        Printer.FontSize = CWidth * 20
        PuzLev = Val(Mid(Printlevnums, PuzNum, 1))
        Call PrintGrid(StartXY(PuzPage, 0), StartXY(PuzPage, 1), CWidth, PrintSud(PuzNum), PuzLev)
        If PuzPage >= PrintPage Then
            PageTxt = "Page " & Trim(Str(PageNumber))
            PuzPage = 0
            Printer.FontSize = 8
            Printer.FontBold = False
            Printer.CurrentX = CenterPage - Printer.TextWidth(PageTxt) / 2
            Printer.CurrentY = Printer.ScaleHeight - Printer.TextHeight(PageTxt) * 2
            Printer.Print PageTxt
            Printer.NewPage
            PageNumber = PageNumber + 1
        End If
    Loop
    If PuzPage <> 1 Then
        PageTxt = "Page " & Trim(Str(PageNumber))
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.CurrentX = CenterPage - Printer.TextWidth(PageTxt)
        Printer.CurrentY = Printer.ScaleHeight - Printer.TextHeight(PageTxt) * 2
        Printer.Print PageTxt
    End If
    Printer.EndDoc
End Sub


Private Sub PrintGrid(Top As Double, Left As Double, Width As Double, Sud As String, Level As Integer)
    Dim x As Integer, y As Integer, A As Integer
    Dim X1 As Double, Y1 As Double
    Dim TXT As String
    Dim TxtMid As Double
    Dim TxtHight As Double
    Dim LevTXT As String
    Dim OldFSize As Double
    Dim From As Double
    Dim Too As Double
    Dim Max As Double
    For y = 0 To 1
        If y = 0 Then Max = Printer.ScaleHeight Else Max = Printer.ScaleWidth
        If y = 0 Then Too = Top + Width * GWidth Else Too = Left + Width * GWidth
        For x = 0 To GWidth
            If y = 0 Then From = Left + Width * x Else From = Top + Width * x
            If Too < Max Then
                If x Mod GridSize = 0 Then Printer.DrawWidth = 8 Else Printer.DrawWidth = 1
                If y = 0 Then
                    Printer.Line (From, Top)-(From, Too), vbBlack
                Else
                    Printer.Line (Left, From)-(Too, From), vbBlack
                End If
            End If
        Next
    Next
    If PrintLev = 1 Then        'print the level names
        OldFSize = Printer.FontSize
        Printer.FontSize = Printer.FontSize / 2.5
        LevTXT = "level : " & LNames(Level)
        TxtHight = Top + Width * GWidth + Printer.TextHeight(LevTXT)
        TxtMid = Left + (Width * GWidth / 2) - Printer.TextWidth(LevTXT) / 2
        Printer.CurrentX = TxtMid
        Printer.CurrentY = TxtHight
        Printer.Print LevTXT
        Printer.FontSize = OldFSize
    End If
    A = 1
    For x = 0 To GWmin1
        For y = 0 To GWmin1
            If Mid(Sud, A, 1) <> "0" Then
                TXT = Mid(Sud, A, 1)
                X1 = Left + y * Width + (Width / 2 - Printer.TextWidth(TXT) / 2)
                Y1 = Top + x * Width + (Width / 2 - Printer.TextHeight(TXT) / 2)
                Printer.CurrentX = X1
                Printer.CurrentY = Y1
                Printer.Print TXT
            End If
            A = A + 1
        Next
    Next
End Sub

