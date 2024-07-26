Attribute VB_Name = "Module1"
Public NEV As String

Public Function ReadInf(InfFile As String, InfTab As String) As String

Dim InfTemp As String

InfTemp = " "

Open InfFile For Input As #2
Do Until EOF(2)
Line Input #2, InfTemp
If InfTab = Left(InfTemp, InStr(InfTemp, "=") - 1) Then
    ReadInf = Mid(InfTemp, InStr(InfTemp, "=") + 1)
End If
Loop
Close #2

End Function

Public Function PrintInf(InfFile As String, InfTab As String, PrintStrs As String) As String

Dim InfTemp As String

InfTemp = " "
Open InfFile & ".bak" For Append As #2
Close #2

Open InfFile For Input As #3
Open InfFile & ".bak" For Output As #2

Do Until EOF(3)
Line Input #3, InfTemp
If InfTab = Left(InfTemp, InStr(InfTemp, "=") - 1) Then
    InfTemp = InfTab & "=" & PrintStrs
End If

Print #2, InfTemp

Loop

Close #3
Close #2

Kill (InfFile)
Name InfFile & ".bak" As InfFile

End Function
