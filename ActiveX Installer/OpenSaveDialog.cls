VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenSaveDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*********************************
' OpenSaveDialog
' Written by Glenn Chittenden Jr.
'--------------------------------------------------
' This code is a combination of examples
' from MSDN, VBNet and vbAccelerator
' as well as my own junk.
'*********************************

' Windows 2000/ME and above
Private Type OFN2
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  pvReserved As Long
  dwReserved As Long
  FlagsEx As Long
End Type

' Windows 9.x/NT
Private Type OFN
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type


Public Enum FLAGCONSTANTS
   OFN_ALLOWMULTISELECT = &H200
   OFN_CREATEPROMPT = &H2000
   OFN_ENABLEHOOK = &H20
   OFN_ENABLESIZING = &H800000
   OFN_ENABLETEMPLATE = &H40
   OFN_ENABLETEMPLATEHANDLE = &H80
   OFN_EXPLORER = &H80000 ' always use this flag
   OFN_EXTENSIONDIFFERENT = &H400
   OFN_FILEMUSTEXIST = &H1000
   OFN_HIDEREADONLY = &H4
   OFN_LONGNAMES = &H200000
   OFN_NOCHANGEDIR = &H8
   OFN_NODEREFERENCELINKS = &H100000
   OFN_NOLONGNAMES = &H40000
   OFN_NONETWORKBUTTON = &H20000
   OFN_NOREADONLYRETURN = &H8000
   OFN_NOTESTFILECREATE = &H10000
   OFN_NOVALIDATE = &H100
   OFN_OVERWRITEPROMPT = &H2
   OFN_PATHMUSTEXIST = &H800
   OFN_READONLY = &H1
   OFN_SHAREAWARE = &H4000
   OFN_SHAREFALLTHROUGH = 2
   OFN_SHARENOWARN = 1
   OFN_SHAREWARN = 0
   OFN_SHOWHELP = &H10
End Enum

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1

' Return Error Values
Private Const CDERR_DIALOGFAILURE = &HFFFF
Private Const CDERR_FINDRESFAILURE = &H6
Private Const CDERR_INITIALIZATION = &H2
Private Const CDERR_LOADRESFAILURE = &H7
Private Const CDERR_LOADSTRFAILURE = &H5
Private Const CDERR_LOCKRESFAILURE = &H8
Private Const CDERR_NOHOOK = &HB
Private Const CDERR_NOTEMPLATE = &H3
Private Const CDERR_MEMLOCKFAILURE = &HA
Private Const CDERR_MEMALLOCFAILURE = &H9
Private Const CDERR_STRUCTSIZE = &H1
Private Const FNERR_BUFFERTOOSMALL = &H3003
Private Const FNERR_INVALIDFILENAME = &H3002
Private Const FNERR_SUBCLASSFAILURE = &H3001

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpOFN As OFN) As Boolean
Private Declare Function GetOpenFileName2 Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpOFN As OFN2) As Boolean
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpOFN As OFN) As Boolean
Private Declare Function GetSaveFileName2 Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpOFN As OFN2) As Boolean

' Local variables
Dim m_Filter As String
Dim m_FilterIndex As Long
Dim m_Filename As String
Dim m_FileTitle As String
Dim m_InitialDir As String
Dim m_DefExt As String
Dim m_DialogTitle As String
Dim m_Flags As FLAGCONSTANTS
Dim m_TemplateName As String
Dim m_FlagsEx As Long
Dim m_Files() As String
Dim m_FileCount As Integer
Dim m_Hook As Boolean

Dim IsME2K As Boolean

Public Event InitDialog(ByVal hDlg As Long)
Public Event FileChange(ByVal hDlg As Long)
Public Event FolderChange(ByVal hDlg As Long)
Public Event DialogOK(ByRef bCancel As Boolean)
Public Event TypeChange(ByVal hDlg As Long)
Public Event DialogClose()
Public Event WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)


Public Property Get Filter() As String
   Filter = m_Filter
End Property

Public Property Let Filter(ByVal f As String)
   m_Filter = f
End Property

Public Property Get FilterIndex() As Long
   FilterIndex = m_FilterIndex
End Property

Public Property Let FilterIndex(ByVal f As Long)
   m_FilterIndex = f
End Property

Public Property Get Filename() As String
   Filename = m_Filename
End Property

Public Property Let Filename(ByVal f As String)
   m_Filename = f
End Property

Public Property Get FileTitle() As String
   FileTitle = m_FileTitle
End Property

Public Property Get InitialDir() As String
   InitialDir = m_InitialDir
End Property

Public Property Let InitialDir(ByVal i As String)
   m_InitialDir = i
End Property

Public Property Get DialogTitle() As String
   DialogTitle = m_DialogTitle
End Property

Public Property Let DialogTitle(ByVal d As String)
   m_DialogTitle = d
End Property

Public Property Get Flags() As FLAGCONSTANTS
   Flags = m_Flags
End Property

Public Property Let Flags(ByVal f As FLAGCONSTANTS)
   m_Flags = f
End Property

Public Property Get DefExt() As String
   DefExt = m_DefExt
End Property

Public Property Let DefExt(ByVal d As String)
   d = Replace(d, ".", "")  ' Just in case
   m_DefExt = d
End Property

Public Property Get TemplateName() As String
   TemplateName = m_TemplateName
End Property

Public Property Let TemplateName(ByVal t As String)
   m_TemplateName = t
End Property

Public Property Get FlagsEx() As Long
   FlagsEx = m_FlagsEx
End Property

Public Property Let FlagsEx(ByVal f As Long)
   m_FlagsEx = f
End Property

Public Property Get FileCount() As Integer
   FileCount = m_FileCount
End Property

Public Property Get Files(ByVal Index As Integer) As String
   Files = m_Files(Index)
End Property

'Public Property Get Hook() As Boolean
'   Hook = m_Hook
'End Property
'
'Public Property Let Hook(ByVal b As Boolean)
'   m_Hook = b
'End Property

'++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function ShowOpen(Optional ByVal OwnerhWnd As Long) As Boolean
   ' Show the dialog and return true if a file was selected.
   Dim uFilter As String, delim As String, fls() As String
   Dim i As Integer, ub As Integer, uFlag As Long
   
   Err.Clear
   
   uFilter = Replace(m_Filter, "|", vbNullChar) & vbNullChar & vbNullChar
   
   If IsME2K Then
      Dim O2 As OFN2
      O2.FlagsEx = m_FlagsEx
      O2.Flags = m_Flags
      O2.hInstance = App.hInstance
      O2.hwndOwner = OwnerhWnd
      O2.lpstrDefExt = m_DefExt
      O2.lpstrFile = m_Filename & String(1024, 0)
      O2.lpstrFileTitle = m_FileTitle & String(80, 0)
      O2.lpstrFilter = uFilter
      O2.lpstrInitialDir = m_InitialDir & vbNullChar
      O2.lpstrTitle = m_DialogTitle & vbNullChar
      O2.lpTemplateName = m_TemplateName
      O2.lStructSize = Len(O2)
      O2.nFilterIndex = m_FilterIndex
      O2.nMaxFile = Len(O2.lpstrFile)
      O2.nMaxFileTitle = Len(O2.lpstrFileTitle)
      'If m_Flags And OFN_ENABLEHOOK Then
      '   O2.lpfnHook = MOpenSave.FAR(AddressOf OpenSaveProc)
      '   Set MOpenSave.m_Dialog = Me
      'End If
      
      If GetOpenFileName2(O2) Then
         ShowOpen = True
         If m_Flags And OFN_ALLOWMULTISELECT Then
            ' check for multiple files
            Erase m_Files()
            m_FileCount = 0
            
            delim = IIf(m_Flags And OFN_EXPLORER, vbNullChar, " ")
            If InStr(1, O2.lpstrFile, delim) = O2.nFileOffset Then
               ' it's multiple files
               fls = Split(O2.lpstrFile, delim)
               
               m_Filename = fls(0)  ' the path
               ub = UBound(fls())
               If delim = " " Then fls(ub) = Left(fls(ub), InStr(1, fls(ub), vbNullChar) - 1)
               
               For i = 1 To ub
                  If fls(i) <> "" Then
                     ReDim Preserve m_Files(m_FileCount)
                     m_Files(m_FileCount) = fls(i)
                     m_FileCount = m_FileCount + 1
                  End If
               Next
            Else
               ' only one selected.
               m_FileCount = 1
               m_Filename = Left(O2.lpstrFile, InStr(1, O2.lpstrFile, vbNullChar) - 1)
               m_FileTitle = Left(O2.lpstrFileTitle, InStr(1, O2.lpstrFileTitle, vbNullChar) - 1)
            End If
         Else
            m_FileCount = 1
            m_Filename = Left(O2.lpstrFile, InStr(1, O2.lpstrFile, vbNullChar) - 1)
            m_FileTitle = Left(O2.lpstrFileTitle, InStr(1, O2.lpstrFileTitle, vbNullChar) - 1)
         End If
         
         m_Flags = O2.Flags
      End If
   Else
      Dim O As OFN
      O.Flags = m_Flags
      O.hInstance = App.hInstance
      O.hwndOwner = OwnerhWnd
      O.lpstrDefExt = m_DefExt
      O.lpstrFile = m_Filename & String(1024, 0)
      O.lpstrFileTitle = m_FileTitle & String(80, 0)
      O.lpstrFilter = uFilter
      O.lpstrInitialDir = m_InitialDir & vbNullChar
      O.lpstrTitle = m_DialogTitle & vbNullChar
      O.lpTemplateName = m_TemplateName
      O.lStructSize = Len(O)
      O.nFilterIndex = m_FilterIndex
      O.nMaxFile = Len(O.lpstrFile)
      O.nMaxFileTitle = Len(O.lpstrFileTitle)
      O.lpfnHook = 0
      
      If GetOpenFileName(O) Then
         ShowOpen = True
         If m_Flags And OFN_ALLOWMULTISELECT Then
            ' check for multiple files
            Erase m_Files()
            m_FileCount = 0
            
            delim = IIf(m_Flags And OFN_EXPLORER, vbNullChar, " ")
            If InStr(1, O.lpstrFile, delim) < O.nFileOffset Then
               ' it's multiple files
               fls = Split(O.lpstrFile, delim)
               
               m_Filename = fls(0)  ' the path
               ub = UBound(fls())
               If delim = " " Then fls(ub) = Left(fls(ub), InStr(1, fls(ub), vbNullChar) - 1)
               
               For i = 1 To ub
                  If fls(i) <> "" Then
                     ReDim Preserve m_Files(m_FileCount)
                     m_Files(m_FileCount) = fls(i)
                     m_FileCount = m_FileCount + 1
                  End If
               Next
            Else
               ' only one selected.
               m_FileCount = 1
               m_Filename = Left(O.lpstrFile, InStr(1, O.lpstrFile, vbNullChar) - 1)
               m_FileTitle = Left(O.lpstrFileTitle, InStr(1, O.lpstrFileTitle, vbNullChar) - 1)
            End If
         Else
            m_FileCount = 1
            m_Filename = Left(O.lpstrFile, InStr(1, O.lpstrFile, vbNullChar) - 1)
            m_FileTitle = Left(O.lpstrFileTitle, InStr(1, O.lpstrFileTitle, vbNullChar) - 1)
         End If
         
         m_Flags = O.Flags
      End If
   End If
   
End Function

Public Function ShowSave(Optional ByVal OwnerhWnd As Long) As Boolean
   ' Show the dialog and return true if a file was selected.
   Dim uFilter As String, delim As String, fls() As String
   Dim i As Integer, ub As Integer
   
   uFilter = Replace(m_Filter, "|", vbNullChar) & vbNullChar & vbNullChar
   
   If IsME2K Then
      Dim O2 As OFN2
      O2.FlagsEx = m_FlagsEx
      O2.Flags = m_Flags
      O2.hInstance = App.hInstance
      O2.hwndOwner = OwnerhWnd
      O2.lpstrDefExt = m_DefExt
      O2.lpstrFile = m_Filename & String(1024, 0)
      O2.lpstrFileTitle = m_FileTitle & String(80, 0)
      O2.lpstrFilter = uFilter
      O2.lpstrInitialDir = m_InitialDir & vbNullChar
      O2.lpstrTitle = m_DialogTitle & vbNullChar
      O2.lpTemplateName = m_TemplateName
      O2.lStructSize = Len(O2)
      O2.nFilterIndex = m_FilterIndex
      O2.nMaxFile = Len(O2.lpstrFile)
      O2.nMaxFileTitle = Len(O2.lpstrFileTitle)
      O2.lpfnHook = 0
      
      If GetSaveFileName2(O2) Then
         ShowSave = True
         m_FileCount = 1
         Erase m_Files()
         m_Filename = Left(O2.lpstrFile, InStr(1, O2.lpstrFile, vbNullChar) - 1)
         m_FileTitle = Left(O2.lpstrFileTitle, InStr(1, O2.lpstrFileTitle, vbNullChar) - 1)
         m_Flags = O2.Flags
      End If
   Else
      Dim O As OFN
      O.Flags = m_Flags
      O.hInstance = App.hInstance
      O.hwndOwner = OwnerhWnd
      O.lpstrDefExt = m_DefExt
      O.lpstrFile = m_Filename & String(1024, 0)
      O.lpstrFileTitle = m_FileTitle & String(80, 0)
      O.lpstrFilter = uFilter
      O.lpstrInitialDir = m_InitialDir & vbNullChar
      O.lpstrTitle = m_DialogTitle & vbNullChar
      O.lpTemplateName = m_TemplateName
      O.lStructSize = Len(O)
      O.nFilterIndex = m_FilterIndex
      O.nMaxFile = Len(O.lpstrFile)
      O.nMaxFileTitle = Len(O.lpstrFileTitle)
      O.lpfnHook = 0
      
      If GetSaveFileName(O) Then
         ShowSave = True
         m_FileCount = 1
         Erase m_Files()
         m_Filename = Left(O.lpstrFile, InStr(1, O.lpstrFile, vbNullChar) - 1)
         m_FileTitle = Left(O.lpstrFileTitle, InStr(1, O.lpstrFileTitle, vbNullChar) - 1)
         m_Flags = O.Flags
      End If
   End If
End Function

Private Sub Class_Initialize()
   ' see if it's ME or 2K
   IsME2K = (IsWin2000 Or IsWinME)
End Sub

'******************************************************

Public Sub WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)
    RaiseEvent WMCommand(hDlg, wParam, lParam)
End Sub

Public Sub DialogClose()
    RaiseEvent DialogClose
End Sub

Public Function TypeChange(ByVal hDlg As Long) As Long
    RaiseEvent TypeChange(hDlg)
End Function

Public Function InitDialog(ByVal hDlg As Long) As Long
   RaiseEvent InitDialog(hDlg)
End Function

Public Function FileChange(ByVal hDlg As Long) As Long
   RaiseEvent FileChange(hDlg)
End Function

Public Function FolderChange(ByVal hDlg As Long) As Long
   RaiseEvent FolderChange(hDlg)
End Function

Public Function ConfirmOK() As Boolean
Dim bCancel As Boolean
    bCancel = False
    RaiseEvent DialogOK(bCancel)
    If (bCancel) Then
        ConfirmOK = False
    Else
        ConfirmOK = True
    End If
End Function



'********************************************************

Private Function IsWinME() As Boolean
'returns True if running Windows ME
   #If Win32 Then
      Dim OSV As OSVERSIONINFO
      OSV.dwOSVersionInfoSize = Len(OSV)
      If GetVersionEx(OSV) = 1 Then
        'PlatformId contains a value representing the OS.
        'If VER_PLATFORM_WIN32_WINDOWS and
        'dwVerMajor = 4,  and dwVerMinor > 0, return true
         IsWinME = (OSV.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                   (OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 90)
      End If
   #End If
End Function

Private Function IsWin2000() As Boolean
'returns True if running WinNT2000 (NT5)
   #If Win32 Then
      Dim OSV As OSVERSIONINFO
      OSV.dwOSVersionInfoSize = Len(OSV)
      If GetVersionEx(OSV) = 1 Then
        'PlatformId contains a value representing the OS.
        'If VER_PLATFORM_WIN32_NT and dwVerMajor is 5, return true
         IsWin2000 = (OSV.dwPlatformID = VER_PLATFORM_WIN32_NT) And (OSV.dwMajorVersion = 5)
      End If
   #End If
End Function

