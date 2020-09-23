Attribute VB_Name = "Extractor"
Option Explicit

'****************************
' ActiveX Installer Extractor
'--------------------------------------------
' Written by Glenn Chittenden Jr.
' Date:  1/4/2002
'****************************

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

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub CopyMemoryBytesToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As String, Source As Byte, ByVal Length As Long)


Public Sub Main()
    Dim sysDir As String, ret As Long, bytes() As Byte
    Dim license() As Byte, sLicense As String
    Dim sMsg As String, sProd As String, i As Long
    Dim fi As FileHeader, fInfo As FileInfo, sFile As String
    
    ' get the system folder
    sysDir = Space(260)
    ret = GetSystemDirectory(sysDir, 260)
    sysDir = Left(sysDir, ret)
    
    'extract the header... starting at byte 24577
    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
        Get #1, 24577, fi
        
        ' get the license agreement, if any
        If fi.LicenseSize > 0 Then
            ReDim license(fi.LicenseSize - 1)
            Get #1, , license
        End If
        
        ' remove nulls
        sMsg = Left(fi.sMessage, InStr(1, fi.sMessage, vbNullChar) - 1)
        sProd = Left(fi.sProduct, InStr(1, fi.sProduct, vbNullChar) - 1)
        
        ' ask if they want to install
        If MsgBox(sMsg, vbApplicationModal Or vbYesNo, "Setup") = vbNo Then GoTo StopIt
        
        ' show the license dialog
        If fi.LicenseSize > 0 Then
            sLicense = Space(fi.LicenseSize)
            CopyMemoryBytesToStr sLicense, license(0), fi.LicenseSize
            frmLicense.Text1.Text = sLicense
            frmLicense.Show vbModal
            
            ' exit if canceled
            If frmLicense.Canceled Then GoTo StopIt
        End If
    
        ' get all of the files
        For i = 1 To fi.FileCount
        
            Get #1, , fInfo
            ReDim bytes(fInfo.FileSize - 1)
            Get #1, , bytes

            ' save the file to the system folder
            sFile = "\" & Left(fInfo.sFilename, InStr(1, fInfo.sFilename, vbNullChar) - 1)
            If FileExists(sysDir & sFile) Then Kill (sysDir & sFile)
            
            Open sysDir & sFile For Binary As #3
                Put #3, , bytes
            Close #3
            
            ' register it if needed (silently)
            If fInfo.bReg Then ShellExecute 0, "open", "regsvr32.exe", "/s " & sysDir, "", 1
            
        Next
        
        ' let them know we're done
        MsgBox "Finished installing " & sProd & ".", , "Setup"

StopIt:
    Close #1
    
    End
End Sub

Private Function FileExists(ByVal sFile As String) As Boolean
    ' Return true if the file already exists.
    On Error Resume Next
    FileExists = CBool(FileLen(sFile))
    On Error GoTo 0
End Function

