VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Print puzzles"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4410
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowNames 
      Caption         =   "Show level names"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CheckBox chkDoRandom 
      Caption         =   "Print in random order"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton btnSettings 
      Caption         =   "Printer settings"
      Height          =   495
      Left            =   2880
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox cmbPuzPage 
      Height          =   315
      ItemData        =   "frmPrint.frx":0000
      Left            =   2880
      List            =   "frmPrint.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.ComboBox cmbPuzNum 
      Height          =   315
      ItemData        =   "frmPrint.frx":0004
      Left            =   2880
      List            =   "frmPrint.frx":0006
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2400
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print puzzle level"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox chkPuzzlev 
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkPuzzlev 
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkPuzzlev 
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkPuzzlev 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chkPuzzlev 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkPuzzlev 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkPuzzlev 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Number of puzzles on page"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Number of puzzles to print"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PuzLevNums As String

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim x As Integer, y As Integer
    Dim Puzn1 As Integer, Puzn2 As Integer, PuzNum As Integer
    Dim NuPuz As Integer
    Dim NLev As Integer
    Dim TotPuz(6) As Integer
    PuzLevNums = ""
    PrintPuzNum = Val(cmbPuzNum.Text)
    PrintPage = Val(cmbPuzPage.Text)
    For x = 0 To 6
        If chkPuzzlev(x) = 1 Then
            PuzLevNums = PuzLevNums & Trim(Str(x))
        End If
    Next
    If PuzLevNums = "" Then
        MsgBox "Which levels do i need to print"
        Exit Sub
    End If
    frmCollecting.Show
    DoEvents
    ReDim PrintSud(PrintPuzNum)
    Printlevnums = ""
    Puzn1 = PrintPuzNum / Len(PuzLevNums)
    Puzn2 = PrintPuzNum - Puzn1 * (Len(PuzLevNums) - 1)
    PuzNum = Puzn2
    If chkDoRandom = 1 And chkDoRandom.Enabled = True Then
        For x = 1 To Len(PuzLevNums)
            TotPuz(Val(Mid(PuzLevNums, x, 1))) = PuzNum
            PuzNum = Puzn1
        Next
        For x = 1 To PrintPuzNum
            Puzn1 = Rnd(1) * Len(PuzLevNums) + 1
            If Puzn1 > Len(PuzLevNums) Then Puzn1 = Len(PuzLevNums)
            NLev = Val(Mid(PuzLevNums, Puzn1, 1))
            Do While TotPuz(NLev) = 0
                NLev = (NLev + 1) Mod 7
            Loop
            TotPuz(NLev) = TotPuz(NLev) - 1
            PrintSud(x) = GetField(NLev)
            PrintSud(x) = Scramble_Field(PrintSud(x))
            Printlevnums = Printlevnums & Trim(Str(NLev))
        Next
    Else
        NuPuz = 1
        For x = 1 To Len(PuzLevNums)
            For y = 1 To PuzNum
                NLev = Val(Mid(PuzLevNums, x, 1))
                PrintSud(NuPuz) = GetField(NLev)
                PrintSud(NuPuz) = Scramble_Field(PrintSud(NuPuz))
                Printlevnums = Printlevnums & Trim(Str(NLev))
                NuPuz = NuPuz + 1
            Next
            PuzNum = Puzn1
        Next
    End If
    Unload frmCollecting
    DoEvents
    frmPrinting.Show
    DoEvents
    Call PrintPages
    Unload frmPrinting
    DoEvents
    Unload Me
End Sub

Private Sub btnSettings_Click()
    Dim PrtErr As Boolean
    ShowPrinter Me, PrtErr
End Sub

Private Sub chkPuzzlev_Click(Index As Integer)
    Dim x As Integer
    PuzLevNums = ""
    For x = 0 To 6
        If chkPuzzlev(x) = 1 Then
            PuzLevNums = PuzLevNums & Trim(Str(x))
        End If
    Next
    If Len(PuzLevNums) > 1 Then
        chkDoRandom.Enabled = True
    Else
        chkDoRandom.Enabled = False
    End If
End Sub

Private Sub chkShowNames_Click()
    PrintLev = chkShowNames
End Sub

Private Sub cmbPuzNum_Change()
    If cmbPuzNum = "" Then cmbPuzNum = "1"
End Sub

Private Sub cmbPuzPage_Change()
    If cmbPuzPage = "" Then cmbPuzPage = "1"
End Sub

Private Sub Form_Activate()
    Dim x As Integer
    For x = 0 To 6
        chkPuzzlev(x).Caption = LNames(x)
    Next
    cmbPuzNum.Clear
    For x = 1 To 500
        cmbPuzNum.AddItem x
    Next
    cmbPuzNum = PrintPuzNum
    cmbPuzPage.Clear
    For x = 1 To 2
        cmbPuzPage.AddItem x
    Next
    If GridSize = 3 Then cmbPuzPage.AddItem "4"
    cmbPuzPage = PrintPage
    chkDoRandom = 0
    chkDoRandom.Enabled = False
    chkShowNames = PrintLev
End Sub

