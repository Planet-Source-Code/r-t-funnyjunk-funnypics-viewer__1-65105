Attribute VB_Name = "modWebGet"
' a slight upgrade to GetURL.BAS by B.Cem HANER : technical@cemhaner.com
Option Explicit
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
       
Private Const BUFFER_LEN = 256

Public Function WebGetHTML(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
    
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    
    If hInternet Then
        DoEvents
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        
        Do While lReturn <> 0
            DoEvents
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    iResult = InternetCloseHandle(hInternet)
    WebGetHTML = sData
End Function

Public Sub WebGetBinary(sURL As String, Optional sFile As String)
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer
    Dim hInternet As Long, hSession As Long, lReturn As Long
    Dim iFile As Integer
    
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    
    If hInternet Then
        If sFile = "" Then sFile = Mid(sURL, InStrRev(sURL, "/") + 1)
        iFile = FreeFile
        Open sFile For Binary Access Write As #iFile
        
        DoEvents
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        Put #iFile, , sBuffer
        
        Do While lReturn <> 0
            DoEvents
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            Put #iFile, , sBuffer
        Loop
        
        Close #iFile
    End If
   
    iResult = InternetCloseHandle(hInternet)
End Sub
