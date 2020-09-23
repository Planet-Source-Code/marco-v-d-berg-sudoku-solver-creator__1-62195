Attribute VB_Name = "Mod_Sudoku"
Option Explicit
Public UseGuessing As Boolean       'Use guessing if logic can't solve puzzle
Private UsedUseGues As Boolean
Public Cell() As Integer     'Positions for a row or column or Block
Public DebugOn As Boolean           'set debug on or off
Public Grid() As String           'work grid
Public GridSize As Integer          '3x3 or 4x4
Public UseRightClick As Boolean

Private StartGrid() As String     'Start grid
Private BackGrid() As String      'Backup grid for guessing
Private ResetGrid As String
Private ExtSolved() As Integer    'Buffer for solved positions (0)=number of solved positions
Private SolveMethod As Integer      'Method of solving
Private LastMethod As Integer       'Last used method
Public TPoints As Long
Private BuTPoints As Long
Private DebugText As String         'solving text
Private BuDebTxt As String          'Backup for solving text (used in guessing)
Private HintOn As Boolean           'try to find a hint
Private GotHint As Boolean          'set as hint is found
Private HintTXT As String           'Hint
Private RC(1) As Integer
Private RV As String
Private ToGo As Integer
Private BuToGo As Integer
Private MinPossibilities As Long
Private BUMinPos As Long
Public BString As String
Public GWidth As Integer
Public GWmin1 As Integer
Public GNum As Integer
Public ReDoGrid As Boolean          'repaint the grid is size is changed
Public Const PosString As String = "123456789ABCDEFG"
Public MethUsed(1) As Integer
Public UsedMeth As String

'init this module and it's variables
Public Sub Init_Sudoku()
    Dim x As Integer, y As Integer
    Dim A As Integer, b As Integer
    Dim ST1 As Integer, St2 As Integer
    Dim SP As Integer
    If GridSize = 0 Then GridSize = 3
    GWidth = GridSize * GridSize
    GWmin1 = GWidth - 1
    GNum = GWidth * GWidth
    BString = Left(PosString, GWidth)
    ReDim Cell(GWmin1, GWmin1, 1)
    ReDim Grid(GNum)
    ReDim StartGrid(GNum)
    ReDim BackGrid(GNum)
    ReDim ExtSolved(GNum)
    
    For x = 0 To GWmin1
        For y = 0 To GWmin1
            Cell(x, y, 0) = x * GWidth + 1 + y
        Next
    Next
    ST1 = -1
    For x = 0 To GridSize - 1
        For y = 0 To GridSize - 1
            SP = x * GWidth * GridSize + y * GridSize + 1
            ST1 = ST1 + 1
            St2 = -1
            For A = 0 To GridSize - 1
                For b = 0 To GridSize - 1
                    St2 = St2 + 1
                    Cell(ST1, St2, 1) = SP + b + A * GWidth
                Next
            Next
        Next
    Next
    HintOn = False
End Sub

'init the solver and it's variables
Public Sub Init_Solver()
    Dim x As Integer
    For x = 1 To GNum
        StartGrid(x) = ""
        Grid(x) = BString
    Next
    MethUsed(0) = 0
    MethUsed(1) = 0
    UsedMeth = ""
    TPoints = 0
    DebugText = ""
    HintOn = False
    ToGo = GNum
    ResetGrid = String(GNum, "0")
End Sub

'define the grid for the solver and set the reset positions
Public Sub DefGrid(Field As Integer, Value As String)
    StartGrid(Field) = Value
    Mid(ResetGrid, Field, 1) = Value
End Sub

'reset the grid
Public Function DoReset() As String
    DoReset = ResetGrid
End Function

'show the status
Public Sub Show_DebugText()
    If DebugText = "" Then MsgBox ("Nothing to show"): Exit Sub
    frmStatus.Show
    frmStatus.txtStatus = DebugText
End Sub

'try to get a hint from the program
Public Function Get_Hint() As String
    Dim BuUseGuess As Boolean
    Dim Answer As Integer
    Dim Pos As Integer
    HintOn = True
    GotHint = False
    BuUseGuess = UseGuessing
    UseGuessing = False
    DebugText = ""
    Call SolveSudoku(True)
    If GotHint = False Then
        MsgBox "I'm sorry, but i dont have any hints"
    Else
        Answer = MsgBox(HintTXT & vbCrLf & "Put number in place ?", vbYesNo)
        If Answer = vbYes Then
            Pos = (RC(0) - 1) * GWidth + RC(1)
            Get_Hint = Right("000" & Trim(Str(Pos)), 3) & RV
        End If
    End If
    UseGuessing = BuUseGuess
    Call MakeLog(0, "", 1, "1")     'to reset is for solving
    DebugText = ""
    BuDebTxt = ""
    HintOn = False
    GotHint = False
End Function

'try to solve the puzzle
Public Function SolveSudoku(Optional JustTest As Boolean = False) As Boolean
    Dim TestPos As Integer
    Dim T1 As Integer, TS As Integer, x As Integer
    Dim Guess As Integer
    If JustTest = False Then
        frmSolving.Show
        DoEvents
    End If
    ExtSolved(0) = 0
    Guess = 0
    T1 = 0
    BuDebTxt = ""
    TestPos = 0
    UsedUseGues = False
    MinPossibilities = 1
    Call SolveA                 'set start grid
    Do
        If HintOn = True And GotHint = True Then GoTo EndFunc
        If IsSolved = True Then SolveSudoku = True: Exit Do
        If SolveB = False Then
            If SolveC = False Then
                If SolveD = False Then
                    If SolveE = False Then
                        If SolveF = False Then
                            If SolveG = False Then
                                If SolveH = False Then
                                    If UseGuessing = False Then Exit Do
                                    UsedUseGues = True
                                    If TestPos = GNum + 1 Then Exit Do
                                    SolveMethod = 99
                                    If Guess <> 0 Then Call CopyFromBackup Else Call CopyToBackup: T1 = 0
                                    Guess = 1
                                    'try putting one in and see if it can be solved now
                                    Do
                                        TS = Len(Grid(TestPos))
                                        If TS > 1 Then
                                            x = TestPos
                                            Call MakeLog(99, "", x, Mid(Grid(TestPos), T1 + 1, 1))
                                            Grid(TestPos) = Mid(Grid(TestPos), T1 + 1, 1)
                                            T1 = T1 + 1
                                            If T1 = TS Then
                                                T1 = 1
                                                TestPos = TestPos + 1
                                            End If
                                            Call SetGrid(x)
                                            Exit Do
                                        Else
                                            TestPos = TestPos + 1
                                        End If
                                    Loop While TestPos < GNum + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop
    SolveMethod = 0
    Call MakeLog(0, "", 1, "1")
    If TestPos = GNum + 1 Then
        DebugText = BuDebTxt
        TPoints = BuTPoints
    End If
EndFunc:
    If JustTest = False Then
        Unload frmSolving
        DoEvents
    End If
End Function

'store the methods used for later use by the level checker
Private Sub AddUsedMeth()
    Dim NewVal As String
    If SolveMethod = 99 Then
        NewVal = "Z"
        Exit Sub
    Else
        NewVal = Mid(PosString, SolveMethod, 1)
    End If
    If Right(UsedMeth, 1) <> NewVal Then
        UsedMeth = UsedMeth & NewVal
    End If
End Sub

'check if the puzzle is solved
Private Function IsSolved() As Boolean
    Dim Z As Integer, x As Integer, y As Integer
    Dim T() As Integer
    ReDim T(GWidth)
    Z = 0
    For x = 1 To GNum
        If Len(Grid(x)) > 1 Then Exit Function
        T(InStr(BString, Grid(x))) = T(InStr(BString, Grid(x))) + 1
        If (x) Mod GWidth = 0 Then
            Z = Z + 1
            If T(0) <> 0 Then Exit Function
            For y = 1 To GWidth
                If T(y) <> Z Then Exit Function
            Next
        End If
    Next
    IsSolved = True
End Function

'make a backup of certain variables while guessing
Private Sub CopyToBackup()
    Dim x As Integer
    For x = 1 To GNum
        BackGrid(x) = Grid(x)
    Next
    BuDebTxt = DebugText
    BuTPoints = TPoints
    BuToGo = ToGo
    BUMinPos = MinPossibilities
End Sub

'restore certain variables while guessing
Private Sub CopyFromBackup()
    Dim x As Integer
    For x = 1 To GNum
        Grid(x) = BackGrid(x)
    Next
    DebugText = BuDebTxt
    TPoints = BuTPoints
    ToGo = BuToGo
    MinPossibilities = BUMinPos
End Sub
 
'Store the startgrid into the workgrid
Private Sub SolveA()
    Dim x As Integer
    SolveMethod = 1
    For x = 1 To GNum
        If Len(StartGrid(x)) = 1 And Len(Grid(x)) > 1 Then
            Grid(x) = StartGrid(x)
            Call SetGrid(x)
        End If
    Next
End Sub

'Remove the known position from related row, column and block
Private Sub SetGrid(Position As Integer)
    Dim x As Integer, y As Integer
    Dim Row As Integer, Col As Integer, BL As Integer
    Dim PV As String
    Dim Pos As Integer
    If MethUsed(1) <> 99 Then
        MethUsed(0) = MethUsed(0) + 1
        MethUsed(1) = 99
        Call AddUsedMeth
    End If
    Row = Int((Position - 1) / GWidth)
    Col = Position - 1 - Row * GWidth
    BL = Int(Col / GridSize) + Int(Row / GridSize) * GridSize
    'Remove from block
    Call MakeLog(1, "", Position, Grid(Position))
    If GotHint Then Exit Sub
    For x = 0 To GWmin1
        Pos = Cell(BL, x, 1)
        If Len(Grid(Pos)) > 1 Then
            Call RemoveNum(Pos, Position)
            If Len(Grid(Pos)) = 1 Then
                Call MakeLog(1, "B", Pos, Grid(Position))
                If GotHint Then Exit Sub
            End If
        End If
    Next
    'Remove from rows
    For x = 0 To GWmin1
        Pos = Cell(Row, x, 0)
        If Len(Grid(Pos)) > 1 Then
            Call RemoveNum(Pos, Position)
            If Len(Grid(Pos)) = 1 Then
                Call MakeLog(1, "R", Pos, Grid(Position))
                If GotHint Then Exit Sub
            End If
        End If
    Next
    'Remove from columns
    For x = 0 To GWmin1
        Pos = Cell(x, Col, 0)
        If Len(Grid(Pos)) > 1 Then
            Call RemoveNum(Pos, Position)
            If Len(Grid(Pos)) = 1 Then
                Call MakeLog(1, "R", Pos, Grid(Position))
                If GotHint Then Exit Sub
            End If
        End If
    Next
    Do While ExtSolved(0) > 0
        Pos = ExtSolved(ExtSolved(0))
        ExtSolved(0) = ExtSolved(0) - 1
        Call SetGrid(Pos)
    Loop
End Sub

Private Sub RemoveNum(From As Integer, What As Integer)
    Dim Pos As Integer
    Pos = InStr(Grid(From), Grid(What))
    If Pos = 0 Then Exit Sub
    Grid(From) = Left(Grid(From), Pos - 1) & Mid(Grid(From), Pos + 1)
    If Len(Grid(From)) = 1 Then
        ExtSolved(0) = ExtSolved(0) + 1
        ExtSolved(ExtSolved(0)) = From
    End If
End Sub

'Find a number wich appearce ones in a block,row or column
'this number me be on that place
Private Function SolveB() As Boolean
    Dim x As Integer, y As Integer, ST1 As Integer
    Dim A As Integer, b As Integer, C As Integer, D As Integer
    Dim Pos As Integer
    Dim SNum As String
    Dim R1 As Integer, C1 As Integer
    Dim BL1 As String, BL2 As String
    Dim NumUsed() As Integer
    Dim Lg As String                    'Logs
    Dim F1 As Boolean
    ReDim NumUsed(GWidth, 1)
    SolveMethod = 2
    'find a single possebillity in block 3x3
    For x = 0 To GWmin1
        Do
            For y = 0 To GWmin1
                Pos = Cell(x, y, 1)
                GoSub Check_Pos
            Next
            'check if there is a number with 1 appearence
            F1 = False
            For C = 1 To GWidth
                If NumUsed(C, 0) = 1 Then
                    If MethUsed(1) <> SolveMethod Then
                        MethUsed(0) = MethUsed(0) + 1
                        MethUsed(1) = SolveMethod
                        Call AddUsedMeth
                    End If
                    SNum = Mid(BString, C, 1)
                    'Put this number in place and start over
                    Pos = NumUsed(C, 1)
                    D = Cell(x, 0, 1)
                    R1 = Int((D - 1) / GWidth)
                    C1 = D - 1 - R1 * GWidth
                    BL1 = Chr(65 + R1) & Trim(Str(C1 + 1))
                    Call MakeLog(2, "B", Pos, SNum, BL1)
                    SolveB = True
                    If GotHint Then Exit Function
                    Grid(Pos) = SNum
                    Call SetGrid(Pos)
                    F1 = True
                End If
                NumUsed(C, 0) = 0
            Next
        Loop While F1
    Next
    'find a single possebillity in rows
    For x = 0 To GWmin1
        Do
            For y = 0 To GWmin1
                Pos = Cell(x, y, 0)
                GoSub Check_Pos
            Next
            'check if there is a number with 1 appearence
            F1 = False
            For C = 1 To GWidth
                If NumUsed(C, 0) = 1 Then
                    If MethUsed(1) <> SolveMethod Then
                        MethUsed(0) = MethUsed(0) + 1
                        MethUsed(1) = SolveMethod
                        Call AddUsedMeth
                    End If
                    SNum = Mid(BString, C, 1)
                    'Put this number in place and start over
                    Pos = NumUsed(C, 1)
                    Call MakeLog(2, "R", Pos, SNum, "")
                    SolveB = True
                    If GotHint Then Exit Function
                    Grid(Pos) = SNum
                    Call SetGrid(Pos)
                    F1 = True
                End If
                NumUsed(C, 0) = 0
            Next
        Loop While F1
    Next
    'find a single possebillity in columns
    For x = 0 To GWmin1
        Do
            For y = 0 To GWmin1
                Pos = Cell(y, x, 0)
                GoSub Check_Pos
            Next
            'check if there is a number with 1 appearence
            F1 = False
            For C = 1 To GWidth
                If NumUsed(C, 0) = 1 Then
                    If MethUsed(1) <> SolveMethod Then
                        MethUsed(0) = MethUsed(0) + 1
                        MethUsed(1) = SolveMethod
                        Call AddUsedMeth
                    End If
                    SNum = Mid(BString, C, 1)
                    'Put this number in place and start over
                    Pos = NumUsed(C, 1)
                    Call MakeLog(2, "C", Pos, SNum, "")
                    SolveB = True
                    If GotHint Then Exit Function
                    Grid(Pos) = SNum
                    Call SetGrid(Pos)
                    F1 = True
                End If
                NumUsed(C, 0) = 0
            Next
        Loop While F1
    Next

    Exit Function
Check_Pos:
    If Len(Grid(Pos)) = 1 Then Return
    For C = 1 To Len(Grid(Pos))
        D = InStr(BString, Mid(Grid(Pos), C, 1))
        NumUsed(D, 0) = NumUsed(D, 0) + 1
        NumUsed(D, 1) = Pos
    Next
    Return

End Function

'Look in a row or column for a value wich could only appear inside a
'certain block
'if found remove that number from the rest of the rows and columns
'inside that block
Private Function SolveC() As Boolean
    Dim x As Integer, y As Integer, A As Integer, b As Integer, C As Integer
    Dim X1 As Integer, Y1 As Integer
    Dim Pos As Integer, P2 As Integer
    Dim PosVal As String, PV As String, PW As Integer
    Dim R1 As Integer, C1 As Integer, BL1 As String
    Dim CheckVal As String
    Dim Block As Integer
    Dim BL As String
    Dim V() As String
    Dim DL As Boolean
    Dim Lg As String                    'Logs
    SolveMethod = 3
    'check rows for number wich could only appear in a certain block
    For x = 0 To GWmin1
        ReDim V(GWidth)
        For y = 0 To GWmin1
            Pos = Cell(x, y, 0)
            If Len(Grid(Pos)) > 1 Then
                For A = 1 To Len(Grid(Pos))
                    PV = Mid(Grid(Pos), A, 1)
                    PW = InStr(BString, PV)
                    If InStr(V(PW), Trim(Str(Int(y / GridSize)))) = 0 Then
                        V(PW) = V(PW) + Trim(Str(Int(y / GridSize)))
                    End If
                Next
            End If
        Next
        For A = 1 To GWidth
            If Len(V(A)) = 1 Then
                'found number A wich could appear only in row x from block(v(A))
                'define start of block
                Block = Int(x / GridSize) * GridSize + Int(V(A))
                DL = False
                PV = Mid(BString, A, 1)
                For y = 0 To GWmin1
                    Pos = Cell(Block, y, 1)
                    If Int((Pos - 1) / GWidth) <> x Then 'do not check the same row
                        If InStr(Grid(Pos), PV) > 0 Then
                            If MethUsed(1) <> SolveMethod Then
                                MethUsed(0) = MethUsed(0) + 1
                                MethUsed(1) = SolveMethod
                                Call AddUsedMeth
                            End If
                            If DL = False Then
                                Call MakeLog(3, "R", Cell(Block, 0, 1), PV, Chr(65 + x))
                                DL = True
                                SolveC = True
                            End If
                            P2 = InStr(Grid(Pos), PV)
                            Grid(Pos) = Left(Grid(Pos), P2 - 1) & Mid(Grid(Pos), P2 + 1)
                            If Len(Grid(Pos)) = 1 Then
                                Call MakeLog(3, "", Pos, PV)
                                If GotHint Then Exit Function
                                Call SetGrid(Pos)
                            End If
                        End If
                    End If
                Next
            End If
        Next
    Next
                
    'check columns for number wich could only appear in a certain block
    For x = 0 To GWmin1
        ReDim V(GWidth)
        For y = 0 To GWmin1
            Pos = Cell(y, x, 0)
            If Len(Grid(Pos)) > 1 Then
                For A = 1 To Len(Grid(Pos))
                    PV = Mid(Grid(Pos), A, 1)
                    PW = InStr(BString, PV)
                    If InStr(V(PW), Trim(Str(Int(y / GridSize)))) = 0 Then
                        V(PW) = V(PW) + Trim(Str(Int(y / GridSize)))
                    End If
                Next
            End If
        Next
        For A = 1 To GWidth
            If Len(V(A)) = 1 Then
                'found number A wich could appear only in column x from block(v(A))
                'define start of block
                Block = Int(x / GridSize) + Int(V(A)) * GridSize
                DL = False
                PV = Mid(BString, A, 1)
                For y = 0 To GWmin1
                    Pos = Cell(Block, y, 1)
                    If (Pos - 1) Mod GridSize <> x Mod GridSize Then 'do not check the same column
                        If InStr(Grid(Pos), PV) > 0 Then
                            If MethUsed(1) <> SolveMethod Then
                                MethUsed(0) = MethUsed(0) + 1
                                MethUsed(1) = SolveMethod
                                Call AddUsedMeth
                            End If
                            If DL = False Then
                                Call MakeLog(3, "C", Cell(Block, 0, 1), PV, Trim(Str(x + 1)))
                                DL = True
                                SolveC = True
                            End If
                            P2 = InStr(Grid(Pos), PV)
                            Grid(Pos) = Left(Grid(Pos), P2 - 1) & Mid(Grid(Pos), P2 + 1)
                            If Len(Grid(Pos)) = 1 Then
                                Call MakeLog(3, "", Pos, PV)
                                If GotHint Then Exit Function
                                Call SetGrid(Pos)
                            End If
                        End If
                    End If
                Next
            End If
        Next
    Next
End Function

'look in a block for a number wich appear only in a certain row or column
'if found than that number can be savely removed from the rest of the
'row or column outside that block
Private Function SolveD() As Boolean
    Dim x As Integer, y As Integer, A As Integer, b As Integer, Z As Integer
    Dim Sr As Integer, Sc As Integer
    Dim Pos As Integer, PV As String, P1 As Integer, P2 As Integer
    Dim Row As Integer, Col As Integer
    Dim R() As String, C() As String
    Dim DL As Boolean
    Dim BL As String
    Dim Lg As String                    'Logs
    SolveMethod = 4
    For x = 0 To GWmin1
        ReDim R(GWidth)
        ReDim C(GWidth)
        For y = 0 To GWmin1
            Pos = Cell(x, y, 1)
            If Len(Grid(Pos)) > 1 Then
                Row = Int((Pos - 1) / GWidth)
                Col = Pos - 1 - Row * GWidth
                For Z = 1 To Len(Grid(Pos))
                    PV = Mid(Grid(Pos), Z, 1)
                    P2 = InStr(BString, PV)
                    If InStr(R(P2), Trim(Str(Row))) = 0 Then R(P2) = R(P2) & Trim(Str(Row))
                    If InStr(C(P2), Trim(Str(Col))) = 0 Then C(P2) = C(P2) & Trim(Str(Col))
                Next
            End If
        Next
            'check if a number appears only in one row
        For Z = 1 To GWidth
            DL = False
            If Len(R(Z)) = 1 Then
                Row = Val(R(Z))                     'row number
                Sc = (x Mod GridSize)
                PV = Mid(BString, Z, 1)
                For A = 0 To GWmin1
                    If Int(A / GridSize) <> Sc Then
                        Pos = Cell(Row, A, 0)
                        If Len(Grid(Pos)) > 1 Then
                            If InStr(Grid(Pos), PV) > 0 Then
                                If MethUsed(1) <> SolveMethod Then
                                    MethUsed(0) = MethUsed(0) + 1
                                    MethUsed(1) = SolveMethod
                                    Call AddUsedMeth
                                End If
                                If DL = False Then
                                    Call MakeLog(4, "R", Cell(x, 0, 1), PV, Chr(65 + Row))
                                    DL = True
                                    SolveD = True
                                End If
                                P1 = InStr(Grid(Pos), PV)
                                Grid(Pos) = Left(Grid(Pos), P1 - 1) & Mid(Grid(Pos), P1 + 1)
                                Call MakeLog(4, "", Pos, PV)
                                If Len(Grid(Pos)) = 1 Then
                                    If GotHint Then Exit Function
                                    Call SetGrid(Pos)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
                
                
        'check if a number appears only in one column
        For Z = 1 To GWidth
            DL = False
            If Len(C(Z)) = 1 Then
                Col = Val(C(Z))                     'column number
                Sc = Int(x / GridSize)
                PV = Mid(BString, Z, 1)
                For A = 0 To GWmin1
                    If Int(A / GridSize) <> Sc Then
                        Pos = Cell(A, Col, 0)
                        If Len(Grid(Pos)) > 1 Then
                            If InStr(Grid(Pos), PV) > 0 Then
                                If MethUsed(1) <> SolveMethod Then
                                    MethUsed(0) = MethUsed(0) + 1
                                    MethUsed(1) = SolveMethod
                                    Call AddUsedMeth
                                End If
                                If DL = False Then
                                    Call MakeLog(4, "C", Cell(x, 0, 1), PV, Trim(Str(Col + 1)))
                                    DL = True
                                    SolveD = True
                                End If
                                P1 = InStr(Grid(Pos), PV)
                                Grid(Pos) = Left(Grid(Pos), P1 - 1) & Mid(Grid(Pos), P1 + 1)
                                Call MakeLog(4, "", Pos, PV)
                                If Len(Grid(Pos)) = 1 Then
                                    If GotHint Then Exit Function
                                    Call SetGrid(Pos)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
    Next
End Function

'If two values appears twice in a row,column or block and these values
'appears also in the same fields than the
'remaining numbers in those field can be savely removed
'the same goes for 3 field bij values of 3 bytes
'etc.etc.
Private Function SolveE() As Boolean
    Dim x As Integer, y As Integer
    Dim A As Integer, b As Integer, C As Integer
    Dim Pos As Integer
    Dim B1() As Integer
    Dim P2 As Integer
    Dim Lg As String                    'Logs
    Dim Cnt As Integer
    Dim V() As Integer              '5 is out the question
    Dim P() As Integer
    Dim P1 As Integer
    Dim PV As String
    Dim Fail As Boolean
    Dim Test As Boolean
    Dim SetTo As String
    Dim BL() As String
    Dim ST() As String
    Dim ST1 As Integer
    Dim Ch(7) As Integer, ChC As Integer
    Dim R1 As Integer, C1 As Integer
    ReDim BL(GWidth)
    ReDim ST(GWidth)
    SolveMethod = 5
    'check rows
    For x = 0 To GWmin1                      'Rows |
        ReDim V(GWidth, 7)
        For y = 0 To GWmin1                  'columns -
            Pos = Cell(y, x, 0)
            For A = 1 To Len(Grid(Pos))
                PV = Mid(Grid(Pos), A, 1)
                P2 = InStr(BString, PV)
                V(P2, 0) = V(P2, 0) + 1
                If V(P2, 0) < 5 Then V(P2, V(P2, 0)) = Pos
            Next
        Next
        GoSub CheckSet
    Next
    'check Columns
    For x = 0 To GWmin1                      'Rows |
        ReDim V(GWidth, 4)
        For y = 0 To GWmin1                  'columns -
            Pos = Cell(x, y, 0)
            For A = 1 To Len(Grid(Pos))
                PV = Mid(Grid(Pos), A, 1)
                P2 = InStr(BString, PV)
                V(P2, 0) = V(P2, 0) + 1
                If V(P2, 0) < 5 Then V(P2, V(P2, 0)) = Pos
            Next
        Next
        GoSub CheckSet
    Next
    'check blocks
    For x = 0 To GWmin1
        ReDim V(GWidth, 4)
        For y = 0 To GWmin1
            Pos = Cell(x, y, 1)
            For A = 1 To Len(Grid(Pos))
                PV = Mid(Grid(Pos), A, 1)
                P2 = InStr(BString, PV)
                V(P2, 0) = V(P2, 0) + 1
                If V(P2, 0) < 5 Then V(P2, V(P2, 0)) = Pos
            Next
        Next
        GoSub CheckSet
    Next
    Exit Function
        
CheckSet:
    Cnt = 2
    Do While Cnt < 5
        ReDim P(GWidth, GWidth)
        P1 = 0
        Fail = False
        ReDim B1(GWidth)
        For b = 1 To GWidth              'Value
            If V(b, 0) = Cnt Then
                If P1 = 0 Then
                    P1 = 1
                    P(P1, 0) = P(P1, 0) + 1
                    P(P1, P(P1, 0)) = b
                Else
                    Test = False
                    For C = 1 To P1
                        If V(b, 1) = V(P(C, P(C, 0)), 1) Then
                            P(C, 0) = P(C, 0) + 1
                            P(C, P(C, 0)) = b
                            Test = True
                        End If
                    Next
                    If Test = False Then
                        P1 = P1 + 1
                        P(P1, 0) = P(P1, 0) + 1
                        P(P1, P(P1, 0)) = b
                    End If
                End If
            End If
        Next
        For y = 1 To P1
            If P(y, 0) = Cnt Then
                For C = 1 To Cnt
                    B1(0) = Cnt
                    B1(C) = P(y, C)
                Next
                P2 = B1(1)
                For b = 2 To Cnt
                    C = B1(b)
                    For A = 1 To Cnt
                        If V(P2, A) <> V(C, A) Then Fail = True 'not the same
                    Next
                    If Fail Then Exit For
                Next
                SetTo = ""
                For A = 1 To Cnt
                    SetTo = SetTo & Mid(BString, B1(A), 1)
                Next
                b = 0
                For A = 1 To Cnt
                    If Grid(V(P2, A)) = SetTo Then b = b + 1
                Next
                If b = Cnt Then Fail = True
                If Not Fail Then        'succes
                    SetTo = ""
                    For A = 1 To Cnt
                        R1 = Int((V(P2, A) - 1) / GWidth) + 1
                        C1 = V(P2, A) - (R1 - 1) * GWidth
                        BL(A) = Chr(64 + R1) & Trim(Str(C1))
                        ST(A) = Mid(BString, B1(A), 1)
                        SetTo = SetTo & Mid(BString, B1(A), 1)
                    Next
                    Lg = "The values " & ST(1)
                    For A = 2 To Cnt
                        Lg = Lg & " and " & ST(A)
                    Next
                    Lg = Lg & " are set to field [" & BL(1) & "]"
                    For A = 2 To Cnt
                        Lg = Lg & " and [" & BL(A) & "]"
                    Next
                    Call MakeLog(5, "", 1, Lg)
                    SolveE = True
                    For A = 1 To Cnt
                        Grid(V(P2, A)) = SetTo
                    Next
                    If MethUsed(1) <> SolveMethod Then
                        MethUsed(0) = MethUsed(0) + 1
                        MethUsed(1) = SolveMethod
                        Call AddUsedMeth
                    End If
                End If
            End If
        Next
        Cnt = Cnt + 1
    Loop
    Return
        
        
        
End Function

'If a value of two bytes can be found in a square block across two
'3x3 blocks than one of the two values must be in that fields
'just put one value into a field en the other 3 will be solved automaticly
'this is a sort of guessing but there is no other way to do this
Private Function SolveF() As Boolean
    Dim x As Integer, y As Integer
    Dim A As Integer, b As Integer, C As Integer
    Dim A1 As Integer, B1 As Integer
    Dim ST1 As Integer
    Dim R1 As Integer, C1 As Integer
    Dim BLU As String, BRD As String
    Dim P1 As Integer, P2 As Integer
    Dim P3 As Integer, P4 As Integer
    Dim Lg As String                    'Logs
    SolveMethod = 6
    For x = 0 To GWmin1
        If ((x + 1) Mod GridSize) > 0 Then
            For y = 0 To GWmin1 - 1
                P1 = x * GWidth + 1 + y
                If Len(Grid(P1)) = 2 Then
                    For A = 1 To GWmin1 - y
                        P2 = P1 + A
                        If Grid(P1) = Grid(P2) Then
                            For b = 1 To GridSize - ((x + 1) Mod GridSize)
                                P3 = P1 + b * GWidth
                                If Grid(P1) = Grid(P3) Then
                                    P4 = P3 + A
                                    If Grid(P1) = Grid(P4) Then
                                        'found a square
                                        If MethUsed(1) <> SolveMethod Then
                                            MethUsed(0) = MethUsed(0) + 1
                                            MethUsed(1) = SolveMethod
                                            Call AddUsedMeth
                                        End If
                                        R1 = Int((P1 - 1) / GWidth) + 1
                                        C1 = P1 - (R1 - 1) * GWidth
                                        BLU = Chr(64 + R1) & Trim(Str(C1))
                                        R1 = Int((P4 - 1) / GWidth) + 1
                                        C1 = P4 - (R1 - 1) * GWidth
                                        BRD = Chr(64 + R1) & Trim(Str(C1))
                                        Call MakeLog(6, "", 1, BLU, BRD)
                                        Grid(P1) = Left(Grid(P1), 1)
                                        SolveF = True
                                        If GotHint = True Then Exit Function
                                        Call SetGrid(P1)
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
    Next
        
    For x = 0 To GWmin1
        If ((x + 1) Mod GridSize) > 0 Then
            For y = 0 To GWmin1 - 1
                P1 = x + 1 + y * GWidth
                If Len(Grid(P1)) = 2 Then
                    For A = 1 To GWmin1 - y
                        P2 = P1 + A * GWidth
                        If Grid(P1) = Grid(P2) Then
                            For b = 1 To GridSize - ((x + 1) Mod GridSize)
                                P3 = P1 + b
                                If Grid(P1) = Grid(P3) Then
                                    P4 = P3 + A * GWidth
                                    If Grid(P1) = Grid(P4) Then
                                        'found a square
                                        If MethUsed(1) <> SolveMethod Then
                                            MethUsed(0) = MethUsed(0) + 1
                                            MethUsed(1) = SolveMethod
                                            Call AddUsedMeth
                                        End If
                                        R1 = Int((P1 - 1) / GWidth) + 1
                                        C1 = P1 - (R1 - 1) * GWidth
                                        BLU = Chr(64 + R1) & Trim(Str(C1))
                                        R1 = Int((P4 - 1) / GWidth) + 1
                                        C1 = P4 - (R1 - 1) * GWidth
                                        BRD = Chr(64 + R1) & Trim(Str(C1))
                                        Call MakeLog(6, "", 1, BLU, BRD)
                                        Grid(P1) = Left(Grid(P1), 1)
                                        SolveF = True
                                        If GotHint = True Then Exit Function
                                        Call SetGrid(P1)
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
    Next
End Function

'find two numbers of two bytes wich are the same in a row, column or block
'These numbers then can be removed from the rest of the row, column or block
Private Function SolveG() As Boolean
    Dim x As Integer, y As Integer
    Dim X1 As Integer, Y1 As Integer
    Dim BL1 As String, BL2 As String, BL3 As String
    Dim A As Integer, b As Integer, C As Integer, D As Integer
    Dim Pos As Integer, PO As Integer
    Dim PV1 As String, PV2 As String
    Dim P(1) As String
    Dim DL As Boolean
    SolveMethod = 7
    'search the rows
    For x = 0 To GWmin1                  'rows
        For A = 0 To GWmin1 - 1
            PV1 = Grid(Cell(x, A, 0))
            If Len(PV1) = 2 Then
                For b = A + 1 To GWmin1
                    PV2 = Grid(Cell(x, b, 0))
                    If PV2 = PV1 Then       'found a double
                        DL = False
                        P(0) = Left(PV1, 1)
                        P(1) = Right(PV1, 1)
                        For C = 0 To GWmin1      'search the row
                            If C <> A And C <> b Then
                                Pos = Cell(x, C, 0)
                                For D = 0 To 1
                                    If InStr(Grid(Pos), P(D)) > 0 Then
                                        If MethUsed(1) <> SolveMethod Then
                                            MethUsed(0) = MethUsed(0) + 1
                                            MethUsed(1) = SolveMethod
                                            Call AddUsedMeth
                                        End If
                                        If DL = False Then
                                            X1 = Int((Cell(x, A, 0) - 1) / GWidth) + 1
                                            Y1 = Cell(x, A, 0) - (X1 - 1) * GWidth
                                            BL1 = Chr(64 + X1) & Trim(Str(Y1))
                                            X1 = Int((Cell(x, b, 0) - 1) / GWidth) + 1
                                            Y1 = Cell(x, b, 0) - (X1 - 1) * GWidth
                                            BL2 = Chr(64 + X1) & Trim(Str(Y1))
                                            Call MakeLog(7, "R", 0, PV1, BL1, BL2)
                                            DL = True
                                        End If
                                        PO = InStr(Grid(Pos), P(D))
                                        Grid(Pos) = Left(Grid(Pos), PO - 1) & Mid(Grid(Pos), PO + 1)
                                        SolveG = True
                                        Call MakeLog(7, "", Pos, P(D))
                                        If Len(Grid(Pos)) = 1 Then
                                            If GotHint Then Exit Function
                                            Call SetGrid(Pos)
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    Next
        
    'search the columns
    For x = 0 To GWmin1                  'rows
        For A = 0 To GWmin1 - 1
            PV1 = Grid(Cell(A, x, 0))
            If Len(PV1) = 2 Then
                For b = A + 1 To GWmin1
                    PV2 = Grid(Cell(b, x, 0))
                    If PV2 = PV1 Then       'found a double
                        DL = False
                        P(0) = Left(PV1, 1)
                        P(1) = Right(PV1, 1)
                        For C = 0 To GWmin1      'search the row
                            If C <> A And C <> b Then
                                Pos = Cell(C, x, 0)
                                For D = 0 To 1
                                    If InStr(Grid(Pos), P(D)) > 0 Then
                                        If MethUsed(1) <> SolveMethod Then
                                            MethUsed(0) = MethUsed(0) + 1
                                            MethUsed(1) = SolveMethod
                                            Call AddUsedMeth
                                        End If
                                        If DL = False Then
                                            X1 = Int((Cell(A, x, 0) - 1) / GWidth) + 1
                                            Y1 = Cell(A, x, 0) - (X1 - 1) * GWidth
                                            BL1 = Chr(64 + X1) & Trim(Str(Y1))
                                            X1 = Int((Cell(b, x, 0) - 1) / GWidth) + 1
                                            Y1 = Cell(b, x, 0) - (X1 - 1) * GWidth
                                            BL2 = Chr(64 + X1) & Trim(Str(Y1))
                                            Call MakeLog(7, "C", 0, PV1, BL1, BL2)
                                            DL = True
                                        End If
                                        PO = InStr(Grid(Pos), P(D))
                                        Grid(Pos) = Left(Grid(Pos), PO - 1) & Mid(Grid(Pos), PO + 1)
                                        SolveG = True
                                        Call MakeLog(7, "", Pos, P(D))
                                        If Len(Grid(Pos)) = 1 Then
                                            If GotHint Then Exit Function
                                            Call SetGrid(Pos)
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    Next
        
    'search the blocks
    For x = 0 To GWmin1                  'rows
        For A = 0 To GWmin1 - 1
            PV1 = Grid(Cell(x, A, 1))
            If Len(PV1) = 2 Then
                For b = A + 1 To GWmin1
                    PV2 = Grid(Cell(x, b, 1))
                    If PV2 = PV1 Then       'found a double
                        DL = False
                        P(0) = Left(PV1, 1)
                        P(1) = Right(PV1, 1)
                        For C = 0 To GWmin1      'search the row
                            If C <> A And C <> b Then
                                Pos = Cell(x, C, 1)
                                For D = 0 To 1
                                    If InStr(Grid(Pos), P(D)) > 0 Then
                                        If MethUsed(1) <> SolveMethod Then
                                            MethUsed(0) = MethUsed(0) + 1
                                            MethUsed(1) = SolveMethod
                                            Call AddUsedMeth
                                        End If
                                        If DL = False Then
                                            X1 = Int((Cell(x, A, 1) - 1) / GWidth) + 1
                                            Y1 = Cell(x, A, 1) - (X1 - 1) * GWidth
                                            BL1 = Chr(64 + X1) & Trim(Str(Y1))
                                            X1 = Int((Cell(x, b, 1) - 1) / GWidth) + 1
                                            Y1 = Cell(x, b, 1) - (X1 - 1) * GWidth
                                            BL2 = Chr(64 + X1) & Trim(Str(Y1))
                                            X1 = Int((Cell(x, 0, 1) - 1) / GWidth) + 1
                                            Y1 = Cell(x, 0, 1) - (X1 - 1) * GWidth
                                            BL3 = Chr(64 + X1) & Trim(Str(Y1))
                                            Call MakeLog(7, "B", 0, PV1, BL1, BL2, BL3)
                                            DL = True
                                        End If
                                        PO = InStr(Grid(Pos), P(D))
                                        Grid(Pos) = Left(Grid(Pos), PO - 1) & Mid(Grid(Pos), PO + 1)
                                        SolveG = True
                                        Call MakeLog(7, "", Pos, P(D))
                                        If Len(Grid(Pos)) = 1 Then
                                            If GotHint Then Exit Function
                                            Call SetGrid(Pos)
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    Next
    
End Function

'Trying the remaining numbers of a cell and check if a cell anywhere else
'on the grid shows the same value after this test
'this cell must then have this value
Public Function SolveH() As Boolean
    Dim BU() As String
    Dim x As Integer, y As Integer
    Dim Lop As Integer, LopT As Integer
    Dim StPos As Integer
    Dim Pos As Integer
    Dim Value As String
    Dim Donext As Boolean
    Dim Found As Boolean
    Dim BUT As String
    Dim BUP As Long
    Dim BUG As Long
    Dim f(1) As String
    ReDim BU(9, GNum)
    For x = 1 To GNum
        BU(0, x) = Grid(x)
    Next
    BUT = DebugText
    BUP = TPoints
    BUG = ToGo
    StPos = 1
DoNextPos:
    Lop = 1
    Do
        Donext = False
        Do While Len(Grid(StPos)) = 1
            StPos = StPos + 1
            If StPos > GNum Then GoTo StopFunc
        Loop
        LopT = Len(Grid(StPos))
        Grid(StPos) = Mid(Grid(StPos), Lop, 1)
        Call SetGrid(StPos)
        If Lop = LopT Then
            StPos = StPos + 1
            Lop = 0
        End If
        If Lop <> 0 Then
            For x = 1 To GNum
                BU(Lop, x) = Grid(x)
                Grid(x) = BU(0, x)
            Next
            DebugText = BUT
            TPoints = BUP
            ToGo = BUG
            Lop = Lop + 1
        Else
            Exit Do
        End If
    Loop
    Value = ""
    GotHint = False         'hint could be set but shouldn't be
    For x = 1 To GNum
        If Len(BU(0, x)) > 1 Then
            Found = True
            For y = 1 To LopT
                If Len(BU(y, x)) > 1 Then
                    Found = False
                    Exit For
                End If
                If Grid(x) <> BU(y, x) Then
                    Found = False
                    Exit For
                End If
            Next
            If Found = True Then
                Value = Grid(x)
                Pos = x
                Exit For
            End If
        End If
    Next
    For x = 1 To GNum
        Grid(x) = BU(0, x)
    Next
    DebugText = BUT
    TPoints = BUP
    ToGo = BUG
    If Value <> "" Then
        SolveMethod = 8
        If MethUsed(1) <> SolveMethod Then
            MethUsed(0) = MethUsed(0) + 1
            MethUsed(1) = SolveMethod
            Call AddUsedMeth
        End If
        x = Int((StPos - 2) / GWidth) + 1
        y = StPos - 1 - (x - 1) * GWidth
        f(1) = Chr(64 + x) & Trim(Str(y))
        Call MakeLog(8, "", Pos, Value, f(1), Grid(StPos - 1))
        SolveH = True
        If GotHint = True Then Exit Function
        Grid(Pos) = Value
        Call SetGrid(Pos)
        Exit Function
    End If
    If StPos <= GNum Then GoTo DoNextPos
    Exit Function

StopFunc:
    DebugText = BUT
    TPoints = BUP
    ToGo = BUG
End Function

'create the log file
Private Sub MakeLog(SM As Integer, RCB As String, Position As Integer, Optional Ext1 As String, Optional Ext2 As String, Optional Ext3 As String, Optional Ext4 As String)
    Dim Row As Integer, Col As Integer
    Dim x As Integer
    Dim RowL As String
    Dim BL As String, BLlu As String, BLrd As String
    Dim Num As String
    Dim TXT As String
    If LastMethod <> SolveMethod And SolveMethod <> 0 Then
        If LastMethod <> 0 Then Call Print_Contents
        TXT = TXT & vbCrLf
        If SM <> 99 Then
            TXT = TXT & "Solving using method " & SolveMethod
        Else
            TXT = TXT & "Solving using guessing"
        End If
        TXT = TXT & vbCrLf
    End If
    Row = Int((Position - 1) / GWidth) + 1
    RowL = Chr(64 + Row)
    Col = Position - (Row - 1) * GWidth
    BL = RowL & Trim(Str(Col))
    Num = Grid(Position)
    Select Case SM
    Case 0
        TXT = TXT & vbCrLf
        If IsSolved = False Or UsedUseGues = True Then
            TXT = TXT & "minimum possible solution is " & Trim(Str(MinPossibilities))
        Else
            TXT = TXT & "number of possible solution is " & Trim(Str(MinPossibilities))
        End If
        TXT = TXT & vbCrLf
        TXT = TXT & "Last status"
        DebugText = DebugText & TXT & vbCrLf
        TXT = ""
        Call Print_Contents
    Case 1
        Select Case UCase(RCB)
        Case ""
            ToGo = ToGo - 1
            TXT = TXT & "Set cell [" & BL & "] to " & Num & ": Removing " & Num & " from related row, column & block (" & Trim(Str(ToGo)) & " to go)"
        Case "R"
            TXT = TXT & "Removing " & Ext1 & " from [" & BL & "] leaves only " & Num
            TPoints = TPoints + 1
        Case "C"
            TXT = TXT & "Removing " & Ext1 & " from [" & BL & "] leaves only " & Num
            TPoints = TPoints + 1
        Case "B"
            TXT = TXT & "Removing " & Ext1 & " from [" & BL & "] leaves only " & Num
            TPoints = TPoints + 1
        End Select
        If HintOn = True And Len(StartGrid(Position)) = 0 Then
            Select Case UCase(RCB)
            Case "R", "C", "B"
                RC(0) = Row
                RC(1) = Col
                RV = Num
                HintTXT = "Row " & Row & " column " & Col & " could only be " & Num
                GotHint = True
            End Select
        End If
    Case 2
        TPoints = TPoints + 30
        Select Case UCase(RCB)
        Case ""
        Case "R"
            TXT = TXT & "Number " & Ext1 & " appearce only once at [" & BL & "] in row " & RowL
        Case "C"
            TXT = TXT & "Number " & Ext1 & " appearce only once at [" & BL & "] in column " & Trim(Str(Col))
        Case "B"
            TXT = TXT & "Number " & Ext1 & " appearce only once at [" & BL & "] in block [" & Ext2 & "]"
        End Select
        If HintOn = True And Len(StartGrid(Position)) = 0 Then
            Select Case UCase(RCB)
            Case "R", "C", "B"
                RC(0) = Row
                RC(1) = Col
                RV = Ext1
                HintTXT = "Row " & Row & " column " & Col & " could only be " & Ext1
                GotHint = True
            End Select
        End If
    Case 3
        Select Case UCase(RCB)
        Case ""
            TXT = TXT & "Removing " & Ext1 & " from [" & BL & "] leaves only " & Num
        Case "R"
            TXT = TXT & "Number " & Ext1 & " could only exist in block [" & BL & "] in row " & Ext2
            TPoints = TPoints + 100
        Case "C"
            TXT = TXT & "Number " & Ext1 & " could only exist in block [" & BL & "] in column " & Ext2
            TPoints = TPoints + 100
        Case "B"
        End Select
        If HintOn = True And Len(StartGrid(Position)) = 0 Then
            Select Case UCase(RCB)
            Case ""
                RC(0) = Row
                RC(1) = Col
                RV = Num
                HintTXT = "Row " & Row & " column " & Col & " could only be " & Num
                GotHint = True
            End Select
        End If
    Case 4
        TPoints = TPoints + 100
        Select Case UCase(RCB)
        Case ""
            If Len(Num) <> 1 Then
                TXT = TXT & "removing " & Ext1 & " from [" & BL & "] leaves " & Num
            Else
                TXT = TXT & "removing " & Ext1 & " from [" & BL & "] leaves only " & Num
            End If
        Case "R"
            TXT = TXT & "in block [" & BL & "] number " & Ext1 & " could only be in row " & Ext2
            TPoints = TPoints + 100
        Case "C"
            TXT = TXT & "in block [" & BL & "] number " & Ext1 & " could only be in column " & Ext2
            TPoints = TPoints + 100
        Case "B"
        End Select
        If HintOn = True And Len(StartGrid(Position)) = 0 Then
            Select Case UCase(RCB)
            Case ""
                RC(0) = Row
                RC(1) = Col
                RC(2) = Num
                HintTXT = "Row " & Row & " column " & Col & " could only be " & Num
                GotHint = True
            End Select
        End If
    Case 5
        TPoints = TPoints + 80
        TXT = TXT & Ext1
    Case 6
        TPoints = TPoints + 50
        TXT = TXT & "Found square from [" & Ext1 & "] to [" & Ext2 & "]"
        MinPossibilities = MinPossibilities * 2
    Case 7
        Select Case UCase(RCB)
        Case ""
            If Len(Num) <> 1 Then
                TXT = TXT & "removing " & Ext1 & " from [" & BL & "] leaves " & Num
            Else
                TXT = TXT & "removing " & Ext1 & " from [" & BL & "] leaves only " & Num
            End If
        Case "R"
            TXT = TXT & "found " & Ext1 & " in cell [" & Ext2 & "] and [" & Ext3 & "] in row " & Left(Ext3, 1)
            TPoints = TPoints + 50
        Case "C"
            TXT = TXT & "found " & Ext1 & " in cell [" & Ext2 & "] and [" & Ext3 & "] in column " & Right(Ext3, 1)
            TPoints = TPoints + 50
        Case "B"
            TXT = TXT & "found " & Ext1 & " in cell [" & Ext2 & "] and [" & Ext3 & "] in Block " & Ext4
            TPoints = TPoints + 50
        End Select
        If HintOn = True And Len(StartGrid(Position)) = 0 Then
            Select Case UCase(RCB)
            Case ""
                If Len(Num) = 1 Then
                    RC(0) = Row
                    RC(1) = Col
                    RC(2) = Num
                    HintTXT = "By ruling out possibilities, row " & Row & " column " & Col & " could only be " & Num
                    GotHint = True
                End If
            End Select
        End If
    Case 8
        TPoints = TPoints + 200
        TXT = TXT & "Putting " & Left(Ext3, 1)
        For x = 2 To Len(Ext3)
            TXT = TXT & " or " & Mid(Ext3, x, 1)
        Next
        TXT = TXT & " in [" & Ext2 & "] gives " & Ext1 & " in [" & BL & "]"
        If HintOn = True Then
            RC(0) = Row
            RC(1) = Col
            RC(2) = Ext1
            HintTXT = "By making certain assumptions in a field it seems that row " & Row & " column " & Col & " could only be " & Ext1
            GotHint = True
        End If
    Case 99
        TXT = TXT & "Guessing " & Ext1 & " at position [" & BL & "]"
    End Select
    If TXT <> "" Then DebugText = DebugText & TXT & vbCrLf
    LastMethod = SolveMethod
End Sub

'show how the grid is at this moment
Public Sub Print_Contents()
    Dim x As Integer, y As Integer
    Dim R() As String
    Dim C() As String
    Dim Pos As Integer
    Dim T As String
    Dim M As Integer
    Dim TXT As String
    Dim MaxWidth As Integer
    Dim Halfwidth As Integer
    ReDim R(GWmin1)
    ReDim C(GWmin1)
    For x = 1 To GNum
        If MaxWidth < Len(Grid(x)) Then MaxWidth = Len(Grid(x))
    Next
    If MaxWidth < 6 Then MaxWidth = 6
    Halfwidth = Int(MaxWidth / 2)
    For x = 0 To GWmin1
        R(x) = BString
        C(x) = BString
    Next
    TXT = vbCrLf
    TXT = TXT & " |"
    For x = 0 To GWmin1
        TXT = TXT & String(Halfwidth, " ") & Trim(Str(x + 1)) & String(MaxWidth - Halfwidth - Len(Trim(Str(x + 1))), " ")
    Next
    TXT = TXT & vbCrLf
    TXT = TXT & String(MaxWidth * GWidth, "-") & vbCrLf
    For x = 0 To GWmin1
        T = Chr(65 + x) & "|" & String(MaxWidth * GWidth, " ")
        For y = 0 To GWmin1
            Pos = x * GWidth + 1 + y
            M = y * MaxWidth + Halfwidth + 3 - Int(Len(Grid(Pos)) / 2)
            Mid(T, M, Len(Grid(Pos))) = Grid(Pos)
        Next
        TXT = TXT & T & vbCrLf
    Next
    DebugText = DebugText & TXT
End Sub

