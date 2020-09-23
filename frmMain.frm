VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planet Source Code Search"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9975
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkZip 
      Caption         =   "Zip Files"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   140
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkText 
      Caption         =   "Text Files"
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   140
      Value           =   1  'Checked
      Width           =   975
   End
   Begin PSCSearch.ctlDownImgForm myWebImg 
      Height          =   3015
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
   End
   Begin VB.ComboBox cboSort 
      Height          =   315
      ItemData        =   "frmMain.frx":57E2
      Left            =   5280
      List            =   "frmMain.frx":57F2
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Sort Key"
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      ItemData        =   "frmMain.frx":582C
      Left            =   3480
      List            =   "frmMain.frx":5853
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Programming Language"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   285
      Left            =   8880
      TabIndex        =   2
      ToolTipText     =   "Search for hippies in green jump suits"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Search Criteria"
      Top             =   120
      Width           =   3255
   End
   Begin VB.ListBox lstLinks 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Results"
      Top             =   480
      Width           =   5775
   End
   Begin VB.TextBox txtSummary 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      ToolTipText     =   "Program Description"
      Top             =   3600
      Width           =   9735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpts 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSpec 
      Caption         =   "Special"
      Begin VB.Menu mnuSpecNew 
         Caption         =   "Newest Entries"
      End
      Begin VB.Menu mnuSpecMonth 
         Caption         =   "Code Of The Month"
      End
      Begin VB.Menu mnuSpecFame 
         Caption         =   "Hall Of Fame"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop"
      Begin VB.Menu mnuPopVisit 
         Caption         =   "Open Project Page In Default Browser"
      End
      Begin VB.Menu mnuPopPreview 
         Caption         =   "View Preview Screenshot"
      End
      Begin VB.Menu mnuPopGet 
         Caption         =   "Download Project"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim sURL As String
Dim sData As String

Private Type PSCEntry
    Type As String
    Name As String
    Link As String
    Image As String
    Rating As Single
    Description As String
End Type
Dim Entry(1 To 100) As PSCEntry

Dim iCntIndex As Integer

Private Sub cmdSearch_Click()
    Dim sSort As String, sCriteria As String
    Dim i As Integer, iBOF As Long, iEOF As Long
    Dim aEntries() As String
    Dim iPreview As Long
    
    lstLinks.Clear
    txtSummary = ""
    myWebImg.HideImage
    
    If sURL = "" Then
        Select Case cboSort.ListIndex
            Case 0
                sSort = "Alphabetical"
            Case 1
                sSort = "DateDescending"
            Case 2
                sSort = "DateAscending"
            Case 3
                sSort = "CountDescending"
        End Select
        
        sCriteria = txtCriteria
        
        sURL = "http://pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp" & _
        "?optSort=" & sSort & _
        "&cmSearch=Search" & _
        "&txtCriteria=" & txtCriteria & _
        "&chkCodeTypeZip=" & IIf(chkZip.Value, "on", "off") & _
        "&chkCodeTypeText=on" & IIf(chkText.Value, "on", "off") & _
        "&blnResetAllVariables=TRUE" & _
        "&txtMaxNumberOfEntriesPerPage=" & IIf(iMaxEntries > 0 And iMaxEntries < 31, iMaxEntries, 10) & _
        "&chkCodeDifficulty=1%2C+2%2C+3%2C+4&lngWId=" & cboLanguage.ItemData(cboLanguage.ListIndex)
        
        Me.Caption = "Results for """ & sCriteria & """ in " & cboLanguage.Text
    End If
    
Me.Caption = "Getting HTML"
    sData = WebGetHTML(sURL)
Me.Caption = "Parsing Out Garbage HTML"
    iBOF = InStrRev(sData, "<!--Main td")
    iEOF = InStrRev(sData, "<!page info>")
    If iBOF < 1 Or iEOF <= iBOF Then
        Me.Caption = "No Results"
    Else
        sData = Mid(sData, iBOF, iEOF - iBOF)
Me.Caption = "Parsing for Entries"
        sData = Replace(sData, "<FONT Size=2 >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", "")
        aEntries = Split(sData, "<!--descrip-->")
        
        For i = 1 To UBound(aEntries)
            sData = aEntries(i)
            iPreview = 0
        ' set link
            iBOF = InStr(1, sData, "/vb/scripts/")
            iEOF = InStr(iBOF, sData, """>")
            Entry(i).Link = "http://www.pscode.com/" & Mid(sData, iBOF, iEOF - iBOF)
        ' set type
            If InStr(iBOF, sData, "_") > 0 Then
                Entry(i).Type = "CodeZip"
            Else
                Entry(i).Type = "ShowCode"
            End If
        ' set name
            iBOF = InStr(iEOF, sData, "alt=""") + 5
            iEOF = InStr(iBOF, sData, """")
            Entry(i).Name = Mid(sData, iBOF, iEOF - iBOF)
            lstLinks.AddItem Entry(i).Name
        ' set description
            iBOF = InStr(iEOF, sData, "<!description>") + 14
            iEOF = InStr(iBOF, sData, "<a href=""/upload")
            If iEOF <= iBOF Then iEOF = InStr(iBOF, sData, "</font>") Else: iPreview = iEOF
            If iEOF <= iBOF Then iEOF = InStr(iBOF, sData, "<HR>")
            Entry(i).Description = Mid(sData, iBOF, iEOF - iBOF)
        ' set screenshot image
            If iPreview > iBOF Then
                iBOF = InStr(iEOF, sData, "/upload_PSC/")
                iEOF = InStr(iBOF, sData, """") - 1
                Entry(i).Image = "http://www.pscode.com/" & Trim(Mid(sData, iBOF, iEOF - iBOF))
            End If
            
        Next i
    End If
    
    sURL = ""
End Sub

Private Sub Form_Load()
    cboLanguage.ListIndex = GetSetting("PSCSearch", "Settings", "Language", 0)
    cboSort.ListIndex = GetSetting("PSCSearch", "Settings", "SortBy", 3)
    sAccessCode = GetSetting("PSCSearch", "Settings", "AccessCode", "")
    iMaxEntries = GetSetting("PSCSearch", "Settings", "MaxEntries", 10)
    
    myWebImg.Status = "Preview Area"
    myWebImg.StatusVisible = True
    myWebImg.HideImage
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    SaveSetting "PSCSearch", "Settings", "Language", cboLanguage.ListIndex
    SaveSetting "PSCSearch", "Settings", "SortBy", cboSort.ListIndex
    SaveSetting "PSCSearch", "Settings", "AccessCode", sAccessCode
    SaveSetting "PSCSearch", "Settings", "MaxEntries", iMaxEntries
    
    Kill "tmp.jpg"
End Sub

Private Sub lstLinks_Click()
    If lstLinks.ListCount > 0 Then
        If lstLinks.ListIndex > -1 Then
            txtSummary = "Type: " & Entry(lstLinks.ListIndex + 1).Type & vbNewLine & vbNewLine & _
                         "Description: " & Entry(lstLinks.ListIndex + 1).Description
            
            If Entry(lstLinks.ListIndex + 1).Image <> "" Then
                myWebImg.Status = "Right Click for Preview"
            Else
                myWebImg.Status = "No Preview Available"
            End If
            myWebImg.StatusVisible = True
            myWebImg.HideImage
        End If
    End If
End Sub

Private Sub lstLinks_DblClick()
    mnuPopVisit_Click
End Sub

Private Sub lstLinks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If lstLinks.ListCount > 0 Then
            If lstLinks.ListIndex > -1 Then
                PopupMenu mnuPop
            End If
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    MsgBox "The Planet Source Code Search Frontend" & vbNewLine & vbNewLine & _
    "Coded by Locohozt", vbCritical, "About"
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileOpts_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuPopGet_Click()
    If lstLinks.ListCount > 0 Then
        If lstLinks.ListIndex > -1 Then
            If sAccessCode <> "" Then
                Dim sLink As String, sType As String, sName As String
                Dim iBOF As Long, iEOF As Long
                
                sLink = Entry(lstLinks.ListIndex + 1).Link
                sType = Entry(lstLinks.ListIndex + 1).Type
                sName = Entry(lstLinks.ListIndex + 1).Name
                
                iBOF = InStr(1, sLink, "txtCodeId=") + 10
                iEOF = InStr(iBOF, sLink, "&")
                
                If sType = "CodeZip" Then
                    sName = sName & ".zip"
                    sType = "ShowZip"
                ElseIf sType = "ShowCode" Then
                    sName = sName & ".bas"
                    sType = "ShowCodeAsText"
                End If
                
                sLink = "http://pscode.com/vb/scripts/" & sType & ".asp" & _
                        "?lngWId=1&lngCodeId=" & Mid(sLink, iBOF, iEOF - iBOF) & _
                        "&strZipAccessCode=" & sAccessCode
                InputBox 1, 1, sLink
                'WebGetBinary sLink, sName
            Else
                MsgBox "Access Code required. Set it in the options form." & vbNewLine & _
                        "If you don't understand it: Just visit the project page"
            End If
        End If
    End If
End Sub

Private Sub mnuPopPreview_Click()
    If lstLinks.ListCount > 0 Then
        If lstLinks.ListIndex > -1 Then
            Dim sFile As String
            sFile = Entry(lstLinks.ListIndex + 1).Image
            
            If sFile <> "" Then
                myWebImg.Status = "Retrieving Preview"
                myWebImg.StatusVisible = True
                myWebImg.HideImage
                
                myWebImg.DisplayImage sFile
            Else
                myWebImg.Status = "No Preview Available"
                myWebImg.StatusVisible = True
                myWebImg.HideImage
            End If
        End If
    End If
End Sub

Private Sub mnuPopVisit_Click()
    If lstLinks.ListCount > 0 Then
        If lstLinks.ListIndex > -1 Then
            ShellExecute 0, vbNullString, Entry(lstLinks.ListIndex + 1).Link, vbNullString, vbNullString, vbNormalFocus
        End If
    End If
End Sub

Private Sub mnuSpecFame_Click()
    sURL = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&txtMaxNumberOfEntriesPerPage=10&blnTopCode=True&blnResetAllVariables=TRUE&lngWid=1"
    cmdSearch_Click
End Sub

Private Sub mnuSpecMonth_Click()
    sURL = "http://www.planet-source-code.com/vb/contest/ContestAndLeaderBoard.asp?lngWid=1"
    cmdSearch_Click
End Sub

Private Sub mnuSpecNew_Click()
    sURL = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=1"
    cmdSearch_Click
End Sub
