Attribute VB_Name = "Module1"
Public Function MakeLogFile(ByVal ErrorOccurredPosition As String) As String
Dim sFileName As String
Dim iFileNo As Integer
    If Err.Number <> 0 Then
        sFileName = App.Path & "\ErrorLogFile.txt"
        iFileNo = FreeFile
            Open sFileName For Append As #iFileNo
            Print #iFileNo, "Err Number :" & Err.Number
            Print #iFileNo, "Err Description :" & Err.Description
            Print #iFileNo, "Date :" & Date
            Print #iFileNo, "Time :" & Time
            Print #iFileNo, "ErrorOccurredPosition :" & ErrorOccurredPosition
            Close #iFileNo
    End If
End Function

Public Function TextToEncrypt(ByVal EncryptText As String) As String
On Error GoTo Err_Handler
Dim iCharLen As Integer
Dim sMid As String * 1
Dim sReturnString As String
Dim sTotalValue As String
    
    sReturnString = ""
    For iCharLen = 1 To Len(EncryptText)
        sMid = Mid(EncryptText, iCharLen, 1)
                sReturnString = sReturnString & Chr(Asc(sMid) + 34)
                DoEvents
    Next iCharLen
 TextToEncrypt = sReturnString
Exit Function
Err_Handler:
    MsgBox Err.Description, vbInformation + vbOKOnly, App.Title
    Call MakeLogFile("TextToEncrypt(ByVal EncryptText As String)")
End Function

Public Function TextToDecrypt(ByVal DecryptText As String) As String
On Error GoTo Err_Handler
Dim iCharLen As Integer
Dim sMid As String * 1
Dim sReturnString As String
Dim sTotalValue As String
    
    sReturnString = ""
      
    For iCharLen = 1 To Len(DecryptText)
        sMid = Mid(DecryptText, iCharLen, 1)
                sReturnString = sReturnString & Chr(Asc(sMid) - 34)
                DoEvents
    Next iCharLen
 TextToDecrypt = sReturnString
Exit Function
Err_Handler:
    MsgBox Err.Description, vbInformation + vbOKOnly, App.Title
    Call MakeLogFile("TextToDecrypt(ByVal DecryptText As String)")
End Function
