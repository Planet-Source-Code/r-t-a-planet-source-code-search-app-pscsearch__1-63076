VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboMaxEntries 
      Height          =   315
      ItemData        =   "frmOptions.frx":000C
      Left            =   1920
      List            =   "frmOptions.frx":006A
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Apply"
      Height          =   405
      Index           =   2
      Left            =   4200
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Cancel"
      Height          =   405
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "OK"
      Height          =   405
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtAccess 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "C:\src\VB"
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Entries Per Page"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Access Code"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Download Path"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0 ' OK
            Cmd_Click 2
            Cmd_Click 1
        Case 1 ' Cancel
            Unload Me
        Case 2 ' Apply
            iMaxEntries = cboMaxEntries.Text
            If txtAccess <> "" Then sAccessCode = txtAccess
    End Select
End Sub

Private Sub Form_Load()
    If 1 <= iMaxEntries And iMaxEntries <= 30 Then
        cboMaxEntries.ListIndex = iMaxEntries - 1
    Else
        cboMaxEntries.ListIndex = 9
    End If
End Sub

Private Sub lbl_Click(Index As Integer)
    If Index = 2 Then
        MsgBox "Its a cookie set when you login via your web browser." & vbNewLine & _
                "If that doesn't make sense, don't worry about it. Just use your " & _
                "web browser.", vbInformation, "WTF is ""Access Code""?"
    End If
End Sub

Private Sub txtAccess_Click()
    txtAccess.SelStart = 0
    txtAccess.SelLength = Len(txtAccess)
End Sub
