Attribute VB_Name = "Mod_Creator"
Public LNames(6) As String
Private BookFile As String
Private LastLev(1) As Integer
Public Autocreate As Boolean

'Init this module with it's variables
Public Sub Init_Levels()
    LNames(0) = "Easy"
    LNames(1) = "Mild"
    LNames(2) = "Moderate"
    LNames(3) = "Difficult"
    LNames(4) = "Hard"
    LNames(5) = "Harder"
    LNames(6) = "Hardest"
    BookFile = App.Path & "\Sudokubook.sbk"
    Autocreate = False
End Sub

'Get a puzzle from the book
Public Function GetField(Lev As Integer) As String
    Dim x As Integer, y As Integer, Z As Integer
    Dim f() As Integer          'Field defenition
    Dim Nums() As Integer
    Dim Pos As Integer
    Dim KeyName As String
    Dim Sud As String
    Dim IsNew As Boolean
    Dim LName As String
    ReDim Nums(GWmin1)
    LName = Trim(Str(GridSize)) & LNames(Lev)
    x = NumKeys(BookFile, LName)
    If x = 0 Then
        If Autocreate = True Then
            Do
                Sud = Create_Level(Lev)
            Loop While CheckLevel(Sud, 2) <> Lev  'last check
            Call SaveInBook(Sud, True)
            Sud = Transform(Sud)
            LastLev(0) = Lev
            LastLev(1) = 0
            GoTo Set_Values
        Else
            MsgBox "There are no puzzles for this level"
            GetField = ""
            Exit Function
        End If
    End If
    Pos = Int(Rnd(1) * x + 1)
    If Autocreate = True Then
        If LastLev(0) = Lev And LastLev(1) = Pos Then  'don't generate same field positions twice
            Do
                Sud = Create_Level(Lev)
            Loop While CheckLevel(Sud, 2) <> Lev  'last check
            Call SaveInBook(Sud, True)
            Sud = Transform(Sud)
            LastLev(0) = Lev
            LastLev(1) = 0
            GoTo Set_Values
        End If
    End If
    LastLev(0) = Lev
    LastLev(1) = Pos
    KeyName = GetKey(BookFile, LName, Pos - 1)
    Sud = GetKeyVal(BookFile, LName, KeyName)
'get 1-9 in random order
Set_Values:
    For x = 0 To GWmin1
        Do
            IsNew = True
            Pos = Int(Rnd(1) * GWidth + 1)
            If Pos = GWidth + 1 Then IsNew = False
            For y = 0 To x
                If Nums(y) = Pos Then IsNew = False
            Next
            Nums(x) = Pos
        Loop While IsNew = False
    Next
'Place 1-9 in positions
    For x = 0 To GWmin1
        Sud = Replace(Sud, Chr(72 + x), Mid(BString, Nums(x), 1))
    Next
    Sud = Replace(Sud, "x", "0")
    GetField = Sud
End Function

'scramble a puzzle so it looks like a brand new one
Public Function Scramble_Field(Sud As String) As String
    Dim T As String
    Dim TSud As String
    Dim U As Integer
    Dim x As Integer, y As Integer
    Dim X1 As Integer, Y1 As Integer
    Dim A As Integer, b As Integer
    Dim GS As Integer
    Dim DoBlock As Double
    Dim DoRow As Double
    TSud = Sud
    GS = GridSize
    For U = 1 To 40             'do 40 transformations
        DoBlock = Rnd(1)
        If DoBlock > 0.5 Then DoBlock = 1 Else DoBlock = 0
        DoRow = rnd1
        If DoRow > 0.5 Then DoRow = 1 Else DoRow = 0
        A = Int(Rnd(1) * GS): If A = GS Then A = A - 1
        Do
            b = Int(Rnd(1) * GS): If b = GS Then b = b - 1
        Loop While b = A
        Select Case DoBlock
        Case 0
            X1 = Int(Rnd(1) * GS): If X1 = GS Then X1 = X1 - 1
            X1 = X1 * GS
            If DoRow = 1 Then
                For y = 0 To GS * GS - 1
                    T = Mid(TSud, Cell(X1 + A, y, 0), 1)
                    Mid(TSud, Cell(X1 + A, y, 0), 1) = Mid(TSud, Cell(X1 + b, y, 0), 1)
                    Mid(TSud, Cell(X1 + b, y, 0), 1) = T
                Next
            Else
                For y = 0 To GS * GS - 1
                    T = Mid(TSud, Cell(y, X1 + A, 0), 1)
                    Mid(TSud, Cell(y, X1 + A, 0), 1) = Mid(TSud, Cell(y, X1 + b, 0), 1)
                    Mid(TSud, Cell(y, X1 + b, 0), 1) = T
                Next
            End If
        Case 1
            For x = 0 To GS - 1
                If DoRow = 1 Then
                    X1 = A * GS + x
                    Y1 = b * GS + x
                Else
                    X1 = A + x * GS
                    Y1 = b + x * GS
                End If
                For y = 0 To (GS * GS) - 1
                    T = Mid(TSud, Cell(X1, y, 1), 1)
                    Mid(TSud, Cell(X1, y, 1), 1) = Mid(TSud, Cell(Y1, y, 1), 1)
                    Mid(TSud, Cell(Y1, y, 1), 1) = T
                Next
            Next
        End Select
    Next
    Scramble_Field = TSud
End Function

'save a puzzle in the book
Public Sub SaveInBook(Sud As String, Optional AutoSave As Boolean = False)
    Dim x As Integer, y As Integer, Z As Integer
    Dim f() As Integer          'Field defenition
    Dim Nums() As Integer
    Dim PuzNum As Integer
    Dim Pos As Integer
    Dim KeyName As String
    Dim IsNew As Boolean
    Dim LName As String
    Dim Lev As Integer
    ReDim Nums(GWmin1)
    Lev = TestLevel(Sud)
    LName = Trim(Str(GridSize)) & LNames(Lev)
    PuzNum = NumKeys(BookFile, LName) + 1
    For x = 1 To GWidth
        Sud = Replace(Sud, Mid(BString, x, 1), Chr(71 + x))
    Next
    Sud = Replace(Sud, "0", "x")
    If TestBookForDoubles(Sud, Lev) = True Then
        If AutoSave Then Exit Sub       'do not save doubles
        MsgBox "This puzzle is already in the book"
        Exit Sub
    End If
    Call AddToINI(BookFile, LName, Trim(Str(PuzNum)), Sud)
    If AutoSave = False Then MsgBox "Puzzle saved as " & LNames(Lev)
End Sub

'check the book if the same puzzles if not saved twice
Private Function TestBookForDoubles(Sud As String, Lev As Integer) As Boolean
    Dim x As Integer, y As Integer
    Dim LName As String
    Dim KeyName As String
    Dim BSud As String
    LName = Trim(Str(GridSize)) & LNames(Lev)
    x = NumKeys(BookFile, LName)
    TestBookForDoubles = True
    For y = 1 To x + 1
        KeyName = GetKey(BookFile, LName, y - 1)
        BSud = GetKeyVal(BookFile, LName, KeyName)
        If BSud = Sud Then Exit Function
    Next
    TestBookForDoubles = False
End Function

'check the level of the puzzle
Public Function TestLevel(Sud As String, Optional AutoCheck As Integer = 0) As Long
    Dim TestSud As String
    Dim x As Integer, y As Integer, Z As Long
    Dim NumTest As Integer
    Dim Pts As Long
    Dim Meth As Long
    Dim BuGuess As Boolean
    If AutoCheck = 0 Then frmChecking.Show
    DoEvents
    BuGuess = UseGuessing
    UseGuessing = False
    NumTest = 10    'numbers of tests to get the best possible level
    If AutoCheck = 1 Then NumTest = 1
    For x = 1 To NumTest
        TestSud = Sud
        TestSud = Scramble_Field(TestSud)
        Call Init_Solver        'set start values of sudoku-puzzle
        For y = 1 To GNum
            If Mid(TestSud, y, 1) <> "0" Then
                Call DefGrid(y, Mid(TestSud, y, 1))
            End If
        Next
        If SolveSudoku(True) = False Then
            Pts = Pts + 1000
            Meth = Meth + 100000
        End If
        Pts = Pts + TPoints
        For y = 1 To Len(UsedMeth)
            Z = InStr(PosString, Mid(UsedMeth, y, 1))
            Select Case Z
            Case 0
            Case 1
                Meth = Meth + 1 * y
            Case 2
                Meth = Meth + 30 * y
            Case 3
                Meth = Meth + 100 * y
            Case 4
                Meth = Meth + 100 * y
            Case 5
                Meth = Meth + 80 * y
            Case 6
                Meth = Meth + 100 * y
            Case 7
                Meth = Meth + 100 * y
            Case 8
                Meth = Meth + 600 * y
            Case 9
            Case 10
            Case Else
            End Select
        Next
    Next
    UseGuessing = BuGuess
    Pts = Pts / NumTest
    Meth = Meth / NumTest
    Select Case GridSize
        Case 3
            Select Case (Pts * Meth) / 100
                Case Is < 2
                    TestLevel = 0
                Case Is < 200
                    TestLevel = 1
                Case Is < 2800
                    TestLevel = 2
                Case Is < 6000
                    TestLevel = 3
                Case Is < 29000
                    TestLevel = 4
                Case Is < 70000
                    TestLevel = 5
                Case Else
                    TestLevel = 6
            End Select
        Case 4
            Select Case (Pts * Meth) / 100
                Case Is < 2
                    TestLevel = 0
                Case Is < 250
                    TestLevel = 1
                Case Is < 7000
                    TestLevel = 2
                Case Is < 50000
                    TestLevel = 3
                Case Is < 100000
                    TestLevel = 4
                Case Is < 800000
                    TestLevel = 5
                Case Else
                    TestLevel = 6
            End Select
    End Select
    If AutoCheck = 0 Then Unload frmChecking
End Function

Public Function CheckLevel(Sud As String, Optional AutoCheck As Integer = 0) As Integer
    Dim Lev As Integer
    Lev = TestLevel(Sud, AutoCheck)
    If AutoCheck = 0 Then MsgBox "Puzzle scaled as level " & LNames(Lev)
    CheckLevel = Lev
End Function

'create a new puzzle
Public Function Create_Level(Lev As Integer) As String
    Dim x As Integer, y As Integer, A As Integer
    Dim FGrid(9) As String
    Dim OldVal As String
    Dim Setval As String
    Dim SetPos As Integer
    Dim Numpos() As Integer
    Dim TotPos As Integer
    Dim Temp As Integer
    Dim FullGrid As String
    Dim TestGrid As String
    Dim RemCells As Integer
    Dim MaxCells As Integer
    frmNew.Show
    If GridSize = 3 Then
        FGrid(0) = "147398256682475193539612748378249561915867324264531987796154832853726419421983675"
        FGrid(1) = "876192435529634187413578926945813762382467591167925348654381279238759614791246853"
        FGrid(2) = "283714596971265384654938217412689753739451628865327941546893172128576439397142865"
        FGrid(3) = "734296815685134729291587364157843296423659178869721453346972581512468937978315642"
        FGrid(4) = "462879135185326974397541862928765413743218659516493287271634598839152746654987321"
        FGrid(5) = "256381794931467582478259361195632847823714659764895213617543928542978136389126475"
        FGrid(6) = "645913278781652934293784165516438729974526813832179546358247691167395482429861357"
        FGrid(7) = "631429875475618329928357416857192634164735298392846751216583947749261583583974162"
        FGrid(8) = "576821943431697258928354176394578612762913485815246397247165839189732564653489721"
        FGrid(9) = "725394861914586327836127549148239756679415283253678914481753692392861475567942138"
    Else
        FGrid(0) = "7F3D9B51C8A426EGEBC9F4362DG1578AA6182CGDF75E439B5G4278AE63B9F1CDGA5E83D9B6F2C471D39C4F2B7A18E5G621B4C6E7593G8ADF876F5A1G4EDC9B3215A6G7C23F9DBE483D27BEF8G4CA691598FG3564E12B7DAC4CEBD19A8576G2F3BE8A197CD26F3G54C475A2BF1GE3D86962D1EG839C45AFB7F9G36D45AB871C2E"
        FGrid(1) = "A69C1EDF2G4B78534GB2A69C7358FD1E53874GB2FE1DC9A61EDF5387C6A92B4G69C4EDFA5BG21738GB2569C41837AFED3871GB25ADEF4C69EDFA3871496C52GB9C4GDFA632B5E187B2539C4GE7816ADF871EB2536FDAG49CDFA6871EGC9435B2C4GBFA698523DE712538C4GBD17E96FA71ED25389AF6BGC4FA6971EDB4CG8325"
        FGrid(2) = "684E39CGB1AF25D7G93C1BAF725D4E68FB1A275D84E63CG9D72548E693CG1AFBE694BG3CF71A825DCGB37F1AD82594E6AF718D25694EB3CG5D82964EGB3C71AF4EG9FCB3AD7168253CFBDA715682G94E1AD76582EG94FB3C2568GE94CFB3D71A94CGA3FB15D7E682B3AF51D72E68CG94715DE2684CG9AFB382E6C4G93AFB5D71"
        FGrid(3) = "7168D5B3F4GAC9E239CB4A2618ED57FG2FEDCG19756B4A835AG47E8F32C96DB11BACE3G76D428F9543869DF2B75G1ECAG79E5B68ACF1342DF5D214ACE98376GBDC47265A8G1EFB39BE53GF94D627A81C921A8C3E5FB4DG7668FGB7D193AC254EAG31687D2B9FEC54E625F9CG4A38B1D7CDBF3145GE7692A88479A2EBC1D5G36F"
        FGrid(4) = "1CDBA3EG62F958473G7826FC45EA9DB16A5EB9487C1D32FGF9425D71GB83AEC6ED2A7C538164BFG9B3F49862AGDCE175G7C1FAD49E5B63285896E1GBF732CAD486GD42CA3F9E751BA53CGEBF2471869D9EB7D516CAG8243F241F37895DB6GCEA41E98FA7B3C5DG62C2A51B9ED64GF783DF8G6435E9271BAC7B63CG2D18AF495E"
        FGrid(5) = "8C25F7B4E91A3DG6D7A42369CF5GEB81BF3GDEA14867C92591E6G8C523DB4F7A3DGF715E8A24B6C95B176A3FDGC924E8E64C8D927B3F15AGA2894CGB56E1F7D3GE7B3F4C1598DA621963B28AF74DGE5CFA5819ED6CG273B424CD567GAEB3819F7GDAE41632859CFB63B1C5289DFEAG4745F29GD7B1AC683EC89EABF3G476521D"
        FGrid(6) = "93D2FA56G41C7E8B85B13C7469ED2GFA4EFA9DG2387BC615GC768E1B52AFD34972EFD9A843B61C5GB1G3524CFED798A6C6A57GE38129FDB4D4896BF1ACG5E237389EB5CF274GA16DFA5G276EDB8149C31B47G39DCF6A85E26D2C148A953EB7GFE7C84FD9B653GA215G3DC8B71AF2649E296BA13GEDC45F78AF14E6257G983BDC"
        FGrid(7) = "2DF3567A9CGB81E49A4EDCG153F8276B615823FBE47DA9CGCB7GE49812A653DF568A31B74D29EFGCGFE1C2D4BA3576984C379G5F8E61DBA2D9B2A86EGFC74531E51C7A29DGBF64837G6BF53CA84E1D29F8D91E462753GCBA32A4BD8G691CFE75AE958B13F6D2CG4717G649E5CB8A32FD842F67CD319GBA5EB3CDGFA275E49816"
        FGrid(8) = "2DCF975E38G16A4BB7681ADG452C9EF35A49C3F2D6EB187GE3G1B4689A7FDC524975AEC38BFDG26162AGFB1D53C789E4C1EBG285A49673DF38FD4679EG12C5BAA51678EF2D3G4B9C9G3E512C67B4FDA8D48269ABCF5E31G77FBCDG3419A8E62586573C4AF2D9BG1EGE9A85B671432FCDFB24ED91GC8A57361CD32FG7BE65A489"
        FGrid(9) = "2156EBF8DG7C4A93G98F13C5E64AB7D2DB4A76G9382FE15C37CE2AD4B951FG86BE3G41279586DFCA9625GF83ACDE7B148F7C95AD413B2E6G4A1DBE6C2FG7893564FBDC7A8392G5E1A29358BE1DCG647FCGE864317AF59D2B15D7F29G6BE43CA85DB9C712F468A3GEFCA23D56GEB9184778G4A9EBC21356FDE3618G4F57ADC2B9"
    End If
    x = Rnd(1) * 10
    Do While x > 9
        x = Rnd(1) * 10
    Loop
    FullGrid = FGrid(x)
StartOver:
    Setval = ""
    SetPos = 0
    TestGrid = FullGrid
'Create a easy puzzle
    Do
        Do
            x = Int(Rnd(1) * GNum + 1)
            If x > GNum Then x = GNum
            OldVal = Mid(TestGrid, x, 1)
            If OldVal <> "0" Then Mid(TestGrid, x, 1) = "0"
        Loop While OldVal = "0"
    Loop While CheckLevel(TestGrid, 1) = 0
    Mid(TestGrid, x, 1) = OldVal        'reset last value
'look which fields are still filled in
    ReDim Numpos(GNum)
    TotPos = 0
    For x = 1 To Len(TestGrid)
        If Mid(TestGrid, x, 1) <> "0" Then
            Numpos(TotPos) = x
            TotPos = TotPos + 1
        End If
    Next
    TotPos = TotPos - 1
    'scramble this structure
    For x = 1 To TotPos
        y = Int(Rnd(1) * TotPos + 1)
        If y > TotPos Then y = TotPos
        A = Int(Rnd(1) * TotPos + 1)
        If A > TotPos Then A = TotPos
        Temp = Numpos(A)
        Numpos(A) = Numpos(y)
        Numpos(y) = Temp
    Next
    A = 0
'create more empty fields but still holding the easy level
    Do
        x = Numpos(A)
        OldVal = Mid(TestGrid, x, 1)
        If OldVal <> "0" Then
            Mid(TestGrid, x, 1) = "0"
            y = CheckLevel(TestGrid, 1)
            If y <> 0 Then
                Mid(TestGrid, x, 1) = OldVal
                A = A + 1
            Else
                If Rnd(1) > 0.5 Then
                    Mid(TestGrid, x, 1) = OldVal
                    A = A + 1
                Else
                    Numpos(A) = Numpos(TotPos)
                    TotPos = TotPos - 1
                End If
            End If
            If A > TotPos Then Exit Do
        End If
    Loop
'if a higher level is needed then remove field until level is reached
    If Lev > 0 Then
        A = 0
        Do
            x = Numpos(A)
            OldVal = Mid(TestGrid, x, 1)
            If OldVal <> "0" Then
                Mid(TestGrid, x, 1) = "0"
                y = CheckLevel(TestGrid, 1)
                If y = Lev Then Exit Do
                If y > Lev Then
                    Mid(TestGrid, x, 1) = OldVal
                    A = A + 1
                Else
                    If Rnd(1) > 0.5 Then
                        Mid(TestGrid, x, 1) = OldVal
                        A = A + 1
                    Else
                        Numpos(A) = Numpos(TotPos)
                        TotPos = TotPos - 1
                    End If
                End If
                If A > TotPos Then GoTo StartOver
            End If
        Loop
    End If
    Create_Level = TestGrid
    Unload frmNew
End Function

'Replace numbers for chars for easy store in book
Private Function Transform(Sud As String) As String
    Dim x As Integer
    For x = 0 To GWmin1
        Sud = Replace(Sud, Mid(BString, x + 1, 1), Chr(72 + x))
    Next
    Sud = Replace(Sud, "0", "x")
    Transform = Sud
End Function

'this sub creates 15 puzzles for each catagory
Public Sub AutoCreateBook()
    Dim Lev As Integer
    Dim x As Integer
    Dim LName As String
    Dim NuTot As Integer
    Dim Sud As String
    Lev = 0
    Do
        LName = Trim(Str(GridSize)) & LNames(Lev)
        NuTot = NumKeys(BookFile, LName)
        For x = NuTot + 1 To 15
            Do
                Sud = Create_Level(Lev)
            Loop While CheckLevel(Sud, 2) <> Lev  'last check
            Call SaveInBook(Sud, True)
        Next
        Lev = Lev + 1
    Loop While Lev <= UBound(LNames)
End Sub
