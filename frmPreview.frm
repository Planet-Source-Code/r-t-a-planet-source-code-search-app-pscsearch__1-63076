VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Preview"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9555
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Image imgTmp 
         Height          =   2535
         Left            =   1560
         Top             =   1440
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    picBack.Width = Me.ScaleWidth
    picBack.Height = Me.ScaleHeight
    LogicalSize picBack, imgTmp, 0
End Sub
