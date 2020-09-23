VERSION 5.00
Begin VB.Form FrmOptions 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRightClick 
      Caption         =   "Use right click to show possibilities"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create book"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox chkAutoCreate 
      Caption         =   "Auto create new puzzles if needed"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grid size"
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      Begin VB.OptionButton OptGridSize 
         Caption         =   "4 x 4"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton OptGridSize 
         Caption         =   "3 x 3"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CheckBox chkGeussing 
      Caption         =   "Use guessing while solving"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    Dim OldSize As Integer
    OldSize = GridSize
    GridSize = 3
    If OptGridSize(1).Value = True Then GridSize = 4
    If OldSize <> GridSize Then Call frmSOLVER.RepaintGrid
    Unload Me
End Sub

Private Sub chkAutoCreate_Click()
    Autocreate = chkAutoCreate.Value
End Sub

Private Sub chkGeussing_Click()
    UseGuessing = chkGeussing.Value
End Sub

Private Sub chkRightClick_Click()
    If chkRightClick = 1 Then
        UseRightClick = True
    Else
        UseRightClick = False
    End If
End Sub

Private Sub Command1_Click()
    Dim TXT As String
    TXT = "This option will fill the book until 15 puzzles of every level is reached" & vbCrLf
    TXT = TXT & "it will do that for the active GRIDSIZE" & vbCrLf
    TXT = TXT & "WARNING. This can take a long time (especially by 4x4)" & vbCrLf
    TXT = TXT & "Do you like to continue"
    If MsgBox(TXT, vbYesNo, "Create book") = vbYes Then
        Me.Hide
        DoEvents
        Call AutoCreateBook
        Me.Show
    End If
End Sub

Private Sub Form_Load()
    If UseGuessing Then
        chkGeussing.Value = 1
    Else
        chkGeussing.Value = 0
    End If
    If GridSize = 3 Then
        OptGridSize(0).Value = True
        OptGridSize(1).Value = False
    Else
        OptGridSize(0).Value = False
        OptGridSize(1).Value = True
    End If
    If Autocreate Then
        chkAutoCreate.Value = 1
    Else
        chkAutoCreate.Value = 0
    End If
    If UseRightClick = True Then
        chkRightClick = 1
    Else
        chkRightClick = 0
    End If
End Sub
