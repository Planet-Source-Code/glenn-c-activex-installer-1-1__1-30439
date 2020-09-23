VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompiler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveX Installer"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5040
   Icon            =   "frmCompiler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Page 
      Height          =   3900
      Index           =   1
      Left            =   165
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox chkRegister 
         Caption         =   "Use Regsvr32 to register this component"
         Height          =   210
         Left            =   780
         TabIndex        =   19
         Top             =   3000
         Width           =   3300
      End
      Begin VB.CommandButton cmdFilesClear 
         Caption         =   "&Clear"
         Height          =   345
         Left            =   3375
         TabIndex        =   3
         Top             =   3390
         Width           =   885
      End
      Begin VB.CommandButton cmdFilesRemove 
         Caption         =   "&Remove"
         Height          =   345
         Left            =   1860
         TabIndex        =   2
         Top             =   3390
         Width           =   885
      End
      Begin VB.CommandButton cmdFilesAdd 
         Caption         =   "&Add"
         Height          =   345
         Left            =   345
         TabIndex        =   1
         Top             =   3390
         Width           =   885
      End
      Begin MSComctlLib.ListView lvFiles 
         Height          =   2610
         Left            =   195
         TabIndex        =   0
         Top             =   270
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   4604
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Register"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame Page 
      Height          =   3900
      Index           =   0
      Left            =   165
      TabIndex        =   15
      Top             =   480
      Width           =   4695
      Begin VB.CommandButton cmdBrowse 
         Height          =   330
         Left            =   4005
         Picture         =   "frmCompiler.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3352
         Width           =   360
      End
      Begin VB.TextBox txtEXEName 
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   3375
         Width           =   3615
      End
      Begin VB.TextBox txtProduct 
         Height          =   285
         Left            =   195
         MaxLength       =   63
         TabIndex        =   5
         Top             =   495
         Width           =   4200
      End
      Begin VB.TextBox txtMessage 
         Height          =   915
         Left            =   195
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1185
         Width           =   4200
      End
      Begin VB.TextBox txtLicense 
         Height          =   285
         Left            =   195
         TabIndex        =   9
         Top             =   2580
         Width           =   3615
      End
      Begin VB.CommandButton cmdLicense 
         Height          =   330
         Left            =   4005
         Picture         =   "frmCompiler.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2550
         Width           =   360
      End
      Begin VB.Label Label3 
         Caption         =   "&EXE Filename:"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   3150
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "&Product Name:"
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "&Install Message:"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label4 
         Caption         =   "&License Agreement:"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   2355
         Width           =   1410
      End
   End
   Begin VB.Frame Page 
      Height          =   3900
      Index           =   2
      Left            =   165
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtInstructions 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   3555
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   225
         Width           =   4455
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4410
      Left            =   60
      TabIndex        =   14
      Top             =   105
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7779
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Project"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Instructions"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuProjectNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuProjectOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu zzmnuProjectSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuProjectSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu zzmnuProjectSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectCompile 
         Caption         =   "&Compile"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuProjectTest 
         Caption         =   "&Test EXE"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu zzmnuProjectSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpInstructions 
         Caption         =   "&Instructions"
         Shortcut        =   {F1}
      End
      Begin VB.Menu zzmnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmCompiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************
' ActiveX Installer
'------------------------------------------------------
' Written by Glenn Chittenden Jr.
' Date:  1/4/2002
' Email:  hardrequest@hotmail.com
'**********************************
' Feel free to use and modify this code as
' you see fit.  If you make any improvements
' it would be nice if you would send me a copy.
'**********************************
' 1.0 - allows a single file to be extracted
'         and an optional license aggreement
'         to be displayed.
'-----------------------------------------------------
' 1.1 - Allow multiple files to be extracted.
'         The header and FileInfo had to be
'         changed.
'         A totally new interface was created.
'***********************************


' Info for each file
Private Type FileInfo
    FileSize As Long
    bReg As Byte  ' 1 = use regsvr32
    sFilename As String * 128  ' this was 256, but I wanted to save some space
End Type

' Header for the product
Private Type FileHeader
    sProduct As String * 64
    sMessage As String * 256    ' message displayed in the first messagebox
    LicenseSize As Long
    FileCount As Long
End Type

Dim IsDirty As Boolean  ' true if changes were made
Dim ProjectFile As String  ' full path and filename of the project
Dim CurrentTab As Long  ' actually the index of the current frame

' default message box text (%PRODUCT% is replaced by the Product Name)
Private Const DEFAULT_MESSAGE As String = "This will install %PRODUCT% on your system.  Do you want to continue?"

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'==========================
'==========================

Private Sub chkRegister_Click()
    ' Change the subitem value
    Dim i As Long, newValue As String
    
    newValue = IIf(chkRegister.Value, "Yes", "No")
    
    With lvFiles
        For i = 1 To .ListItems.Count
            If .ListItems(i).Selected Then .ListItems(i).SubItems(1) = newValue
        Next
    End With
    
    IsDirty = True
    mnuProjectSave.Enabled = True
End Sub

Private Sub cmdBrowse_Click()
    ' Get the path and filename of the EXE to create.
    Dim dlg As OpenSaveDialog, sFile As String
    Set dlg = New OpenSaveDialog
    
    dlg.DialogTitle = "Setup Filename"
    dlg.Filter = "Programs (*.exe)|*.exe"
    dlg.Flags = OFN_ENABLESIZING Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
    
    If dlg.ShowSave(Me.hwnd) Then
        sFile = dlg.Filename
        If LCase(Right(sFile, 4)) <> ".exe" Then sFile = sFile & ".exe"
        txtEXEName.Text = sFile
    End If
    
    Set dlg = Nothing
End Sub

Private Sub cmdFilesAdd_Click()
    ' display an open dialog so a file can be selected
    
    Dim dlg As OpenSaveDialog, i As Long, lvItem As ListItem, sPath As String
    Set dlg = New OpenSaveDialog
    
    dlg.DialogTitle = "Add File"
    dlg.Filter = "Components (*.ocx, *.dll)|*.ocx;*.dll|All Files (*.*)|*.*"
    dlg.Flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_ENABLESIZING
    
    If dlg.ShowOpen(Me.hwnd) Then
        If dlg.FileCount > 1 Then
            sPath = dlg.Filename & "\"
            For i = 0 To dlg.FileCount - 1
                Set lvItem = lvFiles.ListItems.Add(, , sPath & dlg.Files(i))
                lvItem.SubItems(1) = "No"  ' Don't register by default
            Next
        Else
            Set lvItem = lvFiles.ListItems.Add(, , dlg.Filename)
            lvItem.SubItems(1) = "No"
        End If
        
        IsDirty = True
        mnuProjectSave.Enabled = True
    End If
    
    Set dlg = Nothing
    Set lvItem = Nothing
End Sub

Private Sub cmdFilesClear_Click()
    lvFiles.ListItems.Clear
    IsDirty = True
    mnuProjectSave.Enabled = True
End Sub

Private Sub cmdFilesRemove_Click()
    ' remove all selected items
    Dim i As Long
    
    With lvFiles
        For i = .ListItems.Count To 1 Step -1
            If .ListItems(i).Selected Then .ListItems.Remove i
        Next
    End With
    
    IsDirty = True
    mnuProjectSave.Enabled = True
End Sub

Private Sub cmdLicense_Click()
    ' look for a license file
    Dim dlg As OpenSaveDialog
    Set dlg = New OpenSaveDialog
    
    dlg.DialogTitle = "License File"
    dlg.Filter = "Text Files (*.txt)|*.txt"
    dlg.Flags = OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    
    If dlg.ShowOpen(Me.hwnd) Then txtLicense.Text = dlg.Filename
    
    Set dlg = Nothing
End Sub


Private Sub Form_Load()
    ' load the (so called) instructions
    Open App.Path & "\instructions.txt" For Input As #1
        txtInstructions.Text = Input(LOF(1), 1)
    Close #1
    
    IsDirty = False
    mnuProjectSave.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If AskSave = False Then Cancel = 1
End Sub

Private Sub lvFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' This will probably never be used, but...
    lvFiles.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' update the register checkbox
    chkRegister.Value = IIf(Item.SubItems(1) = "Yes", 1, 0)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub txtComponent_GotFocus()
    ' highlight text
    SendKeys "{HOME}+{END}"
End Sub

Private Sub mnuHelpInstructions_Click()
    ' no help file... just show the instructions tab
    TabStrip1.Tabs(3).Selected = True
End Sub

Private Sub mnuProjectCompile_Click()
    ' first, make sure the needed info has been filled in
    
    If txtProduct.Text = "" Then
        MsgBox "Please enter the name of the product being installed.", , "Missing Product Name"
        TabStrip1.Tabs(1).Selected = True
        txtProduct.SetFocus
        Exit Sub
    End If
    
    If txtEXEName.Text = "" Then
        MsgBox "Please enter the output EXE filename.", , "Missing EXE Name"
        TabStrip1.Tabs(1).Selected = True
        txtEXEName.SetFocus
        Exit Sub
    End If
    
    If lvFiles.ListItems.Count = 0 Then
        MsgBox "You haven't added any files to install.  You must have" & vbCrLf & "at least one file to install.", , "Nothing to install"
        TabStrip1.Tabs(2).Selected = True
        Exit Sub
    End If
    
    If txtMessage.Text = "" Then txtMessage.Text = Replace(DEFAULT_MESSAGE, "%PRODUCT%", txtProduct.Text)
    
    ' That should take care of the required stuff... let's do it.
    Me.MousePointer = 11
    MakeExe
    Me.MousePointer = 0
    
    If MsgBox("Do you want to test it?", vbYesNo, "Run Setup?") = vbYes Then mnuProjectTest_Click
End Sub

Private Sub mnuProjectNew_Click()
    ' clear the controls
    If AskSave = False Then Exit Sub
    
    txtProduct.Text = ""
    txtLicense.Text = ""
    txtMessage.Text = ""
    txtEXEName.Text = ""
    lvFiles.ListItems.Clear
    ProjectFile = ""
    
    ' switch to the first tab and set focus
    TabStrip1.Tabs(1).Selected = True
    txtProduct.SetFocus
    IsDirty = False
    mnuProjectSave.Enabled = False
    Me.Caption = "ActiveX Installer"
End Sub

Private Sub mnuProjectOpen_Click()
    ' open a previously saved project
    If AskSave = False Then Exit Sub
    
    Dim dlg As OpenSaveDialog
    Set dlg = New OpenSaveDialog
    
    dlg.DialogTitle = "Open Project"
    dlg.Filter = "ActiveX Installer Projects (*.aip)|*.aip"
    dlg.Flags = OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    
    If dlg.ShowOpen(Me.hwnd) Then OpenProject dlg.Filename
    
    Set dlg = Nothing
End Sub

Private Sub mnuProjectSave_Click()
    ' short and sweat
    If ProjectFile = "" Then If GetProjectFile = False Then Exit Sub
    SaveProject
End Sub

Private Sub mnuProjectSaveAs_Click()
    ' too easy
    If GetProjectFile = False Then Exit Sub
    SaveProject
End Sub

Private Sub mnuProjectTest_Click()
    ' run the installation
    If ProjectFile = "" Or txtEXEName.Text = "" Then Exit Sub
    ShellExecute Me.hwnd, "open", txtEXEName.Text, "", "", 1
End Sub

Private Sub TabStrip1_Click()
    ' show the correct frame
    Page(CurrentTab).Visible = False
    CurrentTab = TabStrip1.SelectedItem.Index - 1
    Page(CurrentTab).Visible = True
End Sub

Private Sub txtEXEName_Change()
    IsDirty = True
    mnuProjectSave.Enabled = True
End Sub

Private Sub txtLicense_Change()
    IsDirty = True
    mnuProjectSave.Enabled = True
End Sub

Private Sub txtLicense_GotFocus()
    ' highlight text
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtMessage_Change()
    IsDirty = True
    mnuProjectSave.Enabled = True
End Sub

Private Sub txtMessage_GotFocus()
    ' highlight text
    txtMessage.SelStart = 0
    txtMessage.SelLength = Len(txtMessage.Text)
End Sub

Private Sub txtProduct_Change()
    IsDirty = True
    mnuProjectSave.Enabled = True
End Sub

Private Sub txtProduct_GotFocus()
    ' highlight text
    SendKeys "{HOME}+{END}"
End Sub

'==========================
'==========================

Private Function AskSave() As Boolean
    ' This function displays a message box to ask
    ' if the changes should be saved.
    ' Returns false if Cancel was clicked.
    
    Dim ret As VbMsgBoxResult
    
    If IsDirty Then
        ret = MsgBox("Do you want to save your changes?", vbQuestion + vbYesNoCancel, "Save Changes")
        If ret = vbCancel Then Exit Function  ' the only time False is returned
        
        If ret = vbYes Then mnuProjectSave_Click
    End If
    
    AskSave = True
End Function

Private Function GetProjectFile() As Boolean
    ' This function displays the SaveAs dialog
    ' and returns True if a filename as selected.
    
    Dim dlg As OpenSaveDialog
    Set dlg = New OpenSaveDialog
    
    dlg.DialogTitle = "Save As..."
    dlg.Filter = "ActiveX Installer Projects (*.aip)|*.aip"
    dlg.Flags = OFN_ENABLESIZING Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
    
    If dlg.ShowSave(Me.hwnd) Then
        ProjectFile = dlg.Filename
        If LCase(Right(ProjectFile, 4)) <> ".aip" Then ProjectFile = ProjectFile & ".aip"
        Me.Caption = "ActiveX Installer - " & Mid(ProjectFile, InStrRev(ProjectFile, "\") + 1)
        GetProjectFile = True
    End If
    
    Set dlg = Nothing
End Function

Private Sub OpenProject(ByVal sFile As String)
    ' Gee, I wonder what this sub does :o)
    
    Dim sLines() As String, i As Long, e As Long
    Dim section As String, entry As String
    Dim lvItem As ListItem, code As String
    
    ' first, reset everything
    IsDirty = False  ' make sure AskSave doesn't get called twice
    mnuProjectNew_Click
    
    ' read the file and split it into lines of text
    Open sFile For Binary As #5
        code = String(LOF(5), 0)
        Get #5, , code
    Close #5
    sLines = Split(code, vbNullChar)
    
    ' now move through each line and
    ' separate the sections
    For i = 0 To UBound(sLines)
        If sLines(i) = "" Then Exit For ' this shouldn't happen
        
        e = InStr(1, sLines(i), "=")
        If e > 0 Then
        
            section = Left(sLines(i), e - 1)
            entry = Mid(sLines(i), e + 1)
            
            Select Case section
                Case "Product": txtProduct.Text = entry
                Case "Message": txtMessage.Text = entry
                Case "License": txtLicense.Text = entry
                Case "EXEName": txtEXEName.Text = entry
            End Select
        
        Else
        
            ' If there is no section, assume it is one of the files
            e = InStr(1, sLines(i), "|")
            If e > 0 Then
                Set lvItem = lvFiles.ListItems.Add(, , Left(sLines(i), e - 1))
                lvItem.SubItems(1) = Mid(sLines(i), e + 1)
            End If
            
        End If
        
    Next
    
    ' set project name and caption
    ProjectFile = sFile
    Me.Caption = "ActiveX Installer - " & Mid(sFile, InStrRev(sFile, "\") + 1)
    
    Set lvItem = Nothing
    IsDirty = False
    mnuProjectSave.Enabled = False
    mnuProjectTest.Enabled = FileExists(txtEXEName.Text)
End Sub

Private Sub SaveProject()
    ' make sure there is a filename
    If ProjectFile = "" Then Exit Sub
    
    Dim code As String, Files As String, i As Long
    
    ' assemble the code
    code = "Product=" & txtProduct.Text & vbNullChar
    code = code & "Message=" & txtMessage.Text & vbNullChar
    code = code & "License=" & txtLicense.Text & vbNullChar
    code = code & "EXEName=" & txtEXEName.Text & vbNullChar
    
    ' add the files
    With lvFiles.ListItems
        For i = 1 To .Count
            Files = Files & .Item(i).Text & "|" & .Item(i).SubItems(1) & vbNullChar
        Next
    End With
    
    ' delete existing...
    If FileExists(ProjectFile) Then Kill ProjectFile
    
    ' now save it
    Open ProjectFile For Binary As #5
        Put #5, , code
        Put #5, , Files
    Close #5
    
    ' wow... that was easy :o)
    IsDirty = False
    mnuProjectSave.Enabled = False
End Sub

Private Sub MakeExe()
    ' Append all of the files and header information together.
    
    Dim hdr As FileHeader, fInfo As FileInfo
    Dim license() As Byte, install() As Byte
    Dim i As Long, bytes() As Byte
    
    ' delete it
    If FileExists(txtEXEName.Text) Then Kill txtEXEName.Text
    
    ' set up the header
    hdr.sProduct = txtProduct.Text & vbNullChar
    hdr.sMessage = txtMessage.Text & vbNullChar
    hdr.FileCount = lvFiles.ListItems.Count
    
    ' read the license, is any
    If txtLicense.Text <> "" Then
        Open txtLicense.Text For Binary As #1
            hdr.LicenseSize = LOF(1)
            ReDim license(hdr.LicenseSize - 1)
            Get #1, , license
        Close #1
    End If
    
    ' read the extractor
    ReDim install(24575)  ' we already know the size
    Open App.Path & "\axinstall.exe" For Binary As #1
        Get #1, , install
    Close #1
    
    ' start writting the exe
    Open txtEXEName.Text For Binary As #5
        
        ' first, the installer
        Put #5, , install
        ReDim install(0)  ' free it
        
        ' now the header
        Put #5, , hdr
        
        ' and the license file, if there is one
        If hdr.LicenseSize > 0 Then
            Put #5, , license
            ReDim license(0)  ' free it
        End If
        
        ' now, run through each file
        With lvFiles.ListItems
            For i = 1 To .Count
                
                ' fill out the file info
                fInfo.sFilename = Mid(.Item(i).Text, InStrRev(.Item(i).Text, "\") + 1) & vbNullChar
                fInfo.bReg = IIf(.Item(i).SubItems(1) = "Yes", 1, 0)
                
                Open .Item(i).Text For Binary As #1
                    fInfo.FileSize = LOF(1)
                    ReDim bytes(fInfo.FileSize - 1)
                    Get #1, , bytes
                Close #1
                
                ' save it to the exe
                Put #5, , fInfo
                Put #5, , bytes
                
            Next
        End With
    
    Close #5
    
    mnuProjectTest.Enabled = True
End Sub

Private Function FileExists(ByVal sFile As String) As Boolean
    ' Return true if the file already exists.
    On Error Resume Next
    FileExists = CBool(FileLen(sFile))
    On Error GoTo 0
End Function
