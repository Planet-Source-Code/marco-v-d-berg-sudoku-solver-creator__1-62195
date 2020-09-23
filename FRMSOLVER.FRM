VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmSOLVER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SuDoku Solver/Creater"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   132
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test level"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset grid"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton btnHint 
      Caption         =   "Hint"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton BtnShowStatus 
      Caption         =   "Show status"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton btnOptions 
      Caption         =   "Options"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtVerPos 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "FRMSOLVER.frx":0000
      Top             =   1920
      Width           =   150
   End
   Begin VB.TextBox txtHorPos 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "123456789"
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton BtnCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   1320
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCLEAR 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSOLVE 
      Caption         =   "Solve"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtCELL 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      HideSelection   =   0   'False
      Index           =   0
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   0
      Top             =   360
      Width           =   360
   End
   Begin VB.Label LevLabel 
      Caption         =   "Level : "
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveFile 
         Caption         =   "Save as File"
      End
      Begin VB.Menu mnuSaveBook 
         Caption         =   "Save in Book"
      End
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "Level"
      Visible         =   0   'False
      Begin VB.Menu mnuLev 
         Caption         =   "Easy"
         Index           =   0
      End
      Begin VB.Menu mnuLev 
         Caption         =   "Mild"
         Index           =   1
      End
      Begin VB.Menu mnuLev 
         Caption         =   "Moderate"
         Index           =   2
      End
      Begin VB.Menu mnuLev 
         Caption         =   "Difficult"
         Index           =   3
      End
      Begin VB.Menu mnuLev 
         Caption         =   "Hard"
         Index           =   4
      End
      Begin VB.Menu mnuLev 
         Caption         =   "Harder"
         Index           =   5
      End
      Begin VB.Menu mnuLev 
         Caption         =   "Hardest"
         Index           =   6
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu Popup 
         Caption         =   "0"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu Popup 
         Caption         =   "1"
         Index           =   1
      End
      Begin VB.Menu Popup 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu Popup 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu Popup 
         Caption         =   "4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Popup 
         Caption         =   "5"
         Index           =   5
      End
      Begin VB.Menu Popup 
         Caption         =   "6"
         Index           =   6
      End
      Begin VB.Menu Popup 
         Caption         =   "7"
         Index           =   7
      End
      Begin VB.Menu Popup 
         Caption         =   "8"
         Index           =   8
      End
      Begin VB.Menu Popup 
         Caption         =   "9"
         Index           =   9
      End
      Begin VB.Menu Popup 
         Caption         =   "A"
         Index           =   10
      End
      Begin VB.Menu Popup 
         Caption         =   "B"
         Index           =   11
      End
      Begin VB.Menu Popup 
         Caption         =   "C"
         Index           =   12
      End
      Begin VB.Menu Popup 
         Caption         =   "D"
         Index           =   13
      End
      Begin VB.Menu Popup 
         Caption         =   "E"
         Index           =   14
      End
      Begin VB.Menu Popup 
         Caption         =   "F"
         Index           =   15
      End
      Begin VB.Menu Popup 
         Caption         =   "G"
         Index           =   16
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnuPrintThis 
         Caption         =   "This Puzzle"
      End
      Begin VB.Menu mnuPrintBook 
         Caption         =   "Print more puzzles"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmSOLVER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private TextBoxArray As Object 'A class level dynamic Array
Private HposArray As Object
Private VposArray As Object
Private Possible(2) As String '0 = row  1=column  2=cell
Private LastCell As Integer
Private IsLoading As Boolean

Private Sub BtnCreate_Click()
    PopupMenu mnuLevel
End Sub

Private Sub btnHint_Click()
    Dim x As Integer
    Dim RetVal As String
    Call Init_Solver
    For x = 1 To GNum
        If InStr(BString, txtCELL(x).Text) Then
            Call DefGrid(x, txtCELL(x).Text)
        Else
            txtCELL(x).Text = ""
            txtCELL(x).ForeColor = vbBlack
        End If
    Next
    RetVal = Get_Hint
    If RetVal <> "" Then
        txtCELL(Val(Left(RetVal, 3))) = Right(RetVal, 1)
    End If
End Sub

Private Sub btnLoad_Click()
    Dim LoadFilename As String
    Cdlg.DialogTitle = "Select the file you want to load"
    Cdlg.FileName = ""
    Cdlg.ShowOpen
    LoadFilename = Cdlg.FileName
    Call load_File(LoadFilename)
End Sub

Private Sub btnOptions_Click()
    FrmOptions.Show
End Sub

Private Sub btnPrint_Click()
    PopupMenu mnuPrint
End Sub

Private Sub btnReset_Click()
    Dim x As Integer, y As Integer
    Dim Pos As Integer, Value As String
    Dim Reg As String
    IsLoading = True
    Reg = DoReset
    Call ClearField(False)
    Pos = 1
    For x = 0 To GWmin1
        For y = 0 To GWmin1
            Value = Mid(Reg, Pos, 1)
            Pos = Pos + 1
            txtCELL(x * GWidth + 1 + y) = Value
            If Value <> "0" Then Call DefGrid(x * GWidth + 1 + y, Value)
        Next
    Next
    txtCELL(1).SetFocus
    Call GetAllRCPos
    IsLoading = False
End Sub

Private Sub BtnSave_Click()
    PopupMenu mnuSave
End Sub

Private Sub BtnShowStatus_Click()
    Call Show_DebugText
End Sub

Private Sub btnTest_Click()
    Dim Sud As String
    Dim x As Integer
    For x = 1 To GNum
        If txtCELL(x).Text = "" Then
            Sud = Sud & "0"
        Else
            Sud = Sud & txtCELL(x).Text
        End If
    Next

    x = CheckLevel(Sud)
    LevLabel.Caption = "Level : " & LNames(x)

End Sub

Private Sub Form_Initialize()
    Randomize
    Call Init_Levels
    Call Init_PrintData
    Call Init_Sudoku
    Call Switch_Size
    Call ADD_TextBoxes
    UseRightClick = False
    
End Sub

Private Sub Switch_Size()
    Dim x As Integer
    If GridSize = 3 Then
        txtCELL(0).FontSize = 16
        txtCELL(0).FontBold = False
        txtCELL(0).Width = 32
        txtCELL(0).Height = 32
        frmSOLVER.Height = 7620
        frmSOLVER.Width = 7665
    Else
        txtCELL(0).FontSize = 8
        txtCELL(0).FontBold = True
        txtCELL(0).Width = 24
        txtCELL(0).Height = 24
        frmSOLVER.Height = 10455
        frmSOLVER.Width = 9735
    End If
End Sub

Public Sub RepaintGrid()
    Dim x As Integer
    Dim Cntr As Integer
    IsLoading = True
    Set TextBoxArray = Me.Controls("txtCELL")
    Set HposArray = Me.Controls("txtHorPos")
    Set VposArray = Me.Controls("txtVerPos")
    For x = 1 To TextBoxArray.UBound
        Unload TextBoxArray(x)
    Next
    For x = 1 To HposArray.UBound
        Unload HposArray(x)
    Next
    For x = 1 To VposArray.UBound
        Unload VposArray(x)
    Next
    Call Init_PrintData
    Call Init_Sudoku
    Call Switch_Size
    Call ADD_TextBoxes
    Call GetAllRCPos
    IsLoading = False
End Sub

Private Sub ADD_TextBoxes()

    Dim Xleft As Integer, Ytop As Integer
    Dim x As Integer, y As Integer
    Dim T As Integer
    Dim Cntr As Integer
    Set TextBoxArray = Me.Controls("txtCELL")
    Set HposArray = Me.Controls("txtHorPos")
    Set VposArray = Me.Controls("txtVerPos")
    txtCELL(0).Visible = True
    txtHorPos(0).Visible = True
    txtVerPos(0).Visible = True
    Load TextBoxArray(1)
    txtCELL(1).Left = 104
    txtCELL(1).Width = txtCELL(0).Width
    txtCELL(1).Top = 18
    txtCELL(1).Height = txtCELL(0).Height
    txtCELL(1).Visible = True
    LevLabel.Left = Int((txtCELL(1).Left * 2 + txtCELL(1).Width * GWidth + Int(GWidth / GridSize) * 10) / 2) - (LevLabel.Width / 2)
    LevLabel.Top = 0
    For y = 1 To GWidth: For x = 1 To GWidth
        If Not (x = 1 And y = 1) Then
            Cntr = TextBoxArray.UBound + 1
            Load TextBoxArray(Cntr)
            Xleft = txtCELL(1).Left + txtCELL(1).Width * (x - 1) + Int((x - 1) / GridSize) * 10
            Ytop = txtCELL(1).Top + txtCELL(1).Height * (y - 1) + Int((y - 1) / GridSize) * 10
            txtCELL(Cntr).Move Xleft, Ytop
            txtCELL(Cntr).Visible = True
        End If
    Next x: Next y
    For x = 0 To GWmin1
        T = x * GWidth + GWidth
        Cntr = HposArray.UBound + 1
        Load HposArray(Cntr)
        Xleft = txtCELL(T).Left + txtCELL(T).Width + 10
        Ytop = txtCELL(T).Top + Int(txtCELL(T).Height / 2) - Int(txtHorPos(0).Height / 2)
        txtHorPos(Cntr).Move Xleft, Ytop
        txtHorPos(Cntr).Visible = True
        T = GNum - GWmin1 + x
        Cntr = VposArray.UBound + 1
        Load VposArray(Cntr)
        Xleft = txtCELL(T).Left + Int(txtCELL(T).Width / 2) - Int(txtVerPos(0).Width / 2)
        Ytop = txtCELL(T).Top + txtCELL(T).Height + 10
        txtVerPos(Cntr).Move Xleft, Ytop
        txtVerPos(Cntr).Visible = True
    Next
    txtHorPos(0).Visible = False
    txtVerPos(0).Visible = False
    txtCELL(0).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHelp_Click()
    Dim FileName As String
    Dim Pad As String
    FileName = "SudHelp.rtf"
    Pad = App.Path
    If Right(Pad, 1) <> "\" And Right(Pad, 1) <> "/" Then
        Pad = Pad + "\"
    End If
    FileName = Pad & FileName
    If Dir(FileName) = "" Then
        MsgBox "Helpfile is not availlible"
        Exit Sub
    End If
    frmHelp.rtbHelp.FileName = FileName
    frmHelp.rtbHelp.Locked = True
    frmHelp.Show
End Sub

Private Sub mnuLev_Click(Index As Integer)
    Dim x As Integer, y As Integer
    Dim Pos As Integer, Value As String
    Dim Reg As String
    IsLoading = True
    Call ClearField
    Reg = GetField(Index)
    If Reg = "" Then Exit Sub
    Reg = Scramble_Field(Reg)
    Call Init_Solver        'to initialize the reset function
    Pos = 1
    For x = 0 To GWmin1
        For y = 0 To GWmin1
            Value = Mid(Reg, Pos, 1)
            Pos = Pos + 1
            txtCELL(x * GWidth + 1 + y) = Value
            If Value <> "0" Then Call DefGrid(x * GWidth + 1 + y, Value)
        Next
    Next
    txtCELL(1).SetFocus
    Call GetAllRCPos
    LevLabel.Caption = "Level : " & LNames(Index)
    IsLoading = False


End Sub

Private Sub mnuPrintBook_Click()
    frmPrint.Show
End Sub

Private Sub mnuPrintThis_Click()
    Dim x As Integer
    Dim SData As String
    If HasData = False Then
        MsgBox "There is no Puzzle to print"
        Exit Sub
    End If
    For x = 1 To GNum
        If txtCELL(x).Text = "" Then
            SData = SData & "0"
        Else
            SData = SData & txtCELL(x).Text
        End If
    Next
    PrintPage = 1
    PrintPuzNum = 1
    PrintLev = 1
    Printlevnums = TestLevel(SData, 2)
    ReDim PrintSud(1)
    PrintSud(1) = SData
    frmPrinting.Show
    DoEvents
    Call PrintPages
    Unload frmPrinting
End Sub

Private Sub mnuSaveBook_Click()
    Dim x As Integer
    Dim SData As String
    If HasData = False Then
        MsgBox "There is nothing to be saved"
        Exit Sub
    End If
    For x = 1 To GNum
        If txtCELL(x).Text = "" Then
            SData = SData & "0"
        Else
            SData = SData & txtCELL(x).Text
        End If
    Next
    Call SaveInBook(SData)
End Sub

Private Sub mnuSaveFile_Click()
    Dim FileNr As Integer
    Dim SData As String
    Dim SaveName As String
    Dim Answer As Integer
    Dim x As Integer, y As Integer
'    If HasData = False Then
'        MsgBox "There is nothing to be saved"
'        Exit Sub
'    End If
    For x = 1 To GNum
        If txtCELL(x).Text = "" Then
            SData = SData & "0"
        Else
            SData = SData & txtCELL(x).Text
        End If
    Next
Ask_SaveName:
    SaveName = ""
    Cdlg.DialogTitle = "Type in the name you want to save with"
    Cdlg.FileName = ""
    Cdlg.ShowSave
    SaveName = Cdlg.FileName
    If SaveName = "" Then Exit Sub
    If InStr(SaveName, ".sud") = 0 Then SaveName = SaveName & ".sud"
    If Dir(SaveName, vbNormal) <> "" Then
        Answer = MsgBox("File already exists" & Chr(13) & Chr(13) & "Overwrite", vbCritical + vbYesNo)
        If Answer = vbNo Then
            GoTo Ask_SaveName
        End If
        Kill SaveName
    End If
    FileNr = FreeFile
    Open SaveName For Output As #FileNr
    Print #FileNr, SData
    Close #FileNr
End Sub

Private Function HasData() As Boolean
    Dim x As Integer
    Dim F1 As Boolean, F2 As Boolean
    For x = 1 To GNum
        If txtCELL(x).Text = "" Then
            F1 = True
        Else
            F2 = True
        End If
    Next
    If F1 = True And F2 = True Then HasData = True
End Function

Private Sub Popup_Click(Index As Integer)
    If Popup(Index).Checked = True Then
        txtCELL(LastCell) = ""
    Else
        txtCELL(LastCell) = Mid(BString, Index, 1)
    End If
End Sub

Private Sub txtCELL_Change(Index As Integer)
    Dim Row As Integer, Col As Integer, BL As Integer
    Dim TXT As String
    If InStr(BString, txtCELL(Index).Text) = 0 Or txtCELL(Index).Text = "0" Then
        txtCELL(Index).Text = ""
        txtCELL(Index).ForeColor = vbBlack
    End If
    If txtCELL(Index).Text <> "" Then
        Call GetBpos(Index)
        If InStr(Possible(2), txtCELL(Index).Text) = 0 Then
            TXT = "the digit " & txtCELL(Index).Text & " is not possible here." & vbCrLf
            TXT = TXT & "It must already have been used in this row,column or 3x3 block"
            MsgBox TXT
            txtCELL(Index).Text = ""
            Exit Sub
        End If
    End If
    If Index < 81 Then
        txtCELL(Index + 1).SetFocus
    Else
        txtCELL(1).SetFocus
    End If
    If Not IsLoading Then
        Row = Int((Index - 1) / GWidth)
        Col = Index - 1 - Row * GWidth
        Call GetRCpos(Index)
        txtHorPos(Row + 1) = Possible(0)
        txtVerPos(Col + 1) = Possible(1)
    End If
End Sub

Private Sub txtCELL_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim X1 As Integer
    ' check if right button was pressed ,if so popup menu
    If Button = vbRightButton Then
        txtCELL(Index).Enabled = False 'first disable the textbox so that handle to default popup menu
        txtCELL(Index).Enabled = True  'is not available, then immediately enable it. set focus on it
        txtCELL(Index).SetFocus
        If UseRightClick = False Then
            MsgBox "Right click not activated"
            Exit Sub
        End If
        LastCell = Index
        Call GetBpos(Index)     'possible values in this cell
        Popup(0).Visible = True 'Dummy menu item
        For X1 = 1 To 16
            Popup(X1).Visible = False
            Popup(X1).Checked = False
        Next
        If Len(Possible(2)) = 0 Then
            MsgBox "You must have made an error someware" & vbCrLf & "There are no possebilities here"
            Exit Sub
        End If
        For X1 = 1 To Len(Possible(2))
            Popup(InStr(BString, Mid(Possible(2), X1, 1))).Visible = True
        Next
        Popup(0).Visible = False    'Dummy menu item
        If txtCELL(Index) <> "" Then
            If InStr(BString, txtCELL(Index)) > 0 Then
                Popup(InStr(BString, txtCELL(Index))).Checked = True
            End If
        End If
        PopupMenu mnuPopup
    End If
End Sub

Private Sub cmdSOLVE_Click()
    Dim Algo As Integer
    Dim TXT As String
    Dim x As Integer
    Dim SolveD As Boolean
    Call Init_Solver        'set start values of sudoku-puzzle
    IsLoading = True
    For x = 1 To GNum
        If txtCELL(x).Text <> "" Then
            Call DefGrid(x, txtCELL(x).Text)
        Else
            txtCELL(x).Text = ""
            txtCELL(x).ForeColor = vbBlack
        End If
    Next
    If SolveSudoku = False Then
        TXT = "Could not solve this puzzle" & vbCrLf
        If UseGuessing = False Then
            TXT = TXT & "Push 'reset grid' and try it again with option 'use geussing'"
        Else
            TXT = TXT & "This one is extremely difficult, is impossible or has multiple solutions"
        End If
        MsgBox TXT
    End If
    For x = 1 To GNum
        If Len(Grid(x)) = 1 Then txtCELL(x).Text = Grid(x)
    Next
    Call GetAllRCPos
    If DebugOn = True Then frmStatus.Show
    IsLoading = False
End Sub


Private Sub cmdCLEAR_Click()
    Call ClearField
End Sub

Private Sub ClearField(Optional LevelRes As Boolean = True)
    Dim Cntr As Integer
    IsLoading = True
    Call Init_Solver
    For Cntr = 1 To GNum
        txtCELL(Cntr).Text = ""
        txtCELL(Cntr).ForeColor = vbRed
    Next Cntr
    Call GetAllRCPos
    If LevelRes Then LevLabel.Caption = "Level : "
    IsLoading = False
End Sub


Private Sub load_File(Name As String)
    Dim FreeNum As Integer
    Dim x As Integer, y As Integer
    Dim Pos As Integer
    Dim Reg As String
    Dim Value As String
    If Name = "" Then Exit Sub
    IsLoading = True
    Call ClearField
    FreeNum = FreeFile
    Open Name For Input As #FreeNum
    If InStr(Name, ".sud") = 0 Then
        For x = 1 To GWidth
            Input #FreeNum, Reg
            If Len(Reg) <> GWidth Then
                MsgBox "Illigal puzzle file"
                Exit Sub
            End If
            For y = 1 To GWidth
                If Mid(Reg, (y - 1) * 2 + 1) = 0 Then
                    Value = ""
                Else
                    Value = Mid(Reg, (y - 1) * 2 + 1)
                End If
                txtCELL((x - 1) * GWidth + 1 + (y - 1)) = Value
            Next
        Next
    Else
        Pos = 1
        Input #FreeNum, Reg
        If Len(Reg) <> GNum Then
            MsgBox "Illigal puzzle file"
            Exit Sub
        End If
        For x = 0 To GWmin1
            For y = 0 To GWmin1
                Value = Mid(Reg, Pos, 1)
                Pos = Pos + 1
                txtCELL(x * GWidth + 1 + y) = Value
            Next
        Next
    End If
    Close #FreeNum
    Call GetAllRCPos
    IsLoading = False
End Sub

Private Function GetRCpos(Index As Integer)
    Dim x As Integer
    Dim P2 As Integer
    Dim Row As Integer, Col As Integer, BL As Integer
    Dim V As String
    Row = Int((Index - 1) / GWidth)
    Col = Index - 1 - Row * GWidth
    BL = Int(Col / GridSize) + Int(Row / GridSize) * GridSize
    Possible(0) = BString
    Possible(1) = BString
    For x = 0 To GWmin1
        V = txtCELL(Cell(Row, x, 0))
        If V <> "" Then
            P2 = InStr(Possible(0), V)
            If P2 > 0 Then Possible(0) = Left(Possible(0), P2 - 1) & Mid(Possible(0), P2 + 1)
        End If
        V = txtCELL(Cell(x, Col, 0))
        If V <> "" Then
            P2 = InStr(Possible(1), V)
            If P2 > 0 Then Possible(1) = Left(Possible(1), P2 - 1) & Mid(Possible(1), P2 + 1)
        End If
    Next
End Function

Private Function GetBpos(Index As Integer)
    Dim x As Integer
    Dim P2 As Integer
    Dim Row As Integer, Col As Integer, BL As Integer
    Dim V As String
    Row = Int((Index - 1) / GWidth)
    Col = Index - 1 - Row * GWidth
    BL = Int(Col / GridSize) + Int(Row / GridSize) * GridSize
    Possible(2) = BString
    For x = 0 To GWmin1
        V = txtCELL(Cell(Row, x, 0))
        If V <> "" And Cell(Row, x, 0) <> Index Then
            P2 = InStr(Possible(2), V)
            If P2 > 0 Then Possible(2) = Left(Possible(2), P2 - 1) & Mid(Possible(2), P2 + 1)
        End If
        V = txtCELL(Cell(x, Col, 0))
        If V <> "" And Cell(x, Col, 0) <> Index Then
            P2 = InStr(Possible(2), V)
            If P2 > 0 Then Possible(2) = Left(Possible(2), P2 - 1) & Mid(Possible(2), P2 + 1)
        End If
        V = txtCELL(Cell(BL, x, 1))
        If V <> "" And Cell(BL, x, 1) <> Index Then
            P2 = InStr(Possible(2), V)
            If P2 > 0 Then Possible(2) = Left(Possible(2), P2 - 1) & Mid(Possible(2), P2 + 1)
        End If
    Next
End Function

Private Sub GetAllRCPos()
    Dim x As Integer
    For x = 1 To GWidth
        Call GetRCpos((x - 1) * GWidth + x)
        txtHorPos(x) = Possible(0)
        txtVerPos(x) = Possible(1)
    Next
End Sub
