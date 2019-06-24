Attribute VB_Name = "Utility"
Const REVISION_PROPERTY_NAME = "מהדורה"


Function formatTableString(st As String) As String
' Certain strings in tables give wierd results, this solves these errors
formatTableString = Left(st, Len(st) - 2)
End Function

Function retPropVal(prop As String) As Variant
' Retruns the file's property named prop

retPropVal = "Prop Error"

For Each p In ActiveDocument.CustomDocumentProperties

    If prop = p.Name Then
        retPropVal = p.Value
    End If
Next p

End Function

Sub changeProperties(prop As String, st As String)
' Changes certain property to st

For Each p In ActiveDocument.CustomDocumentProperties
    If prop = p.Name Then
        p.Value = st
    End If
Next p

End Sub

Function URLChange(oldStr As String, newStr As String)
' Changes all URL containing oldStr with newStr

For Each link In ActiveDocument.Hyperlinks
    link.address = Replace(link.address, oldStr, newStr)
Next link

End Function

Function checkURL(strURL As String) As Boolean
' Checks if url is valid

checkURL = False
Dim objDemand As Object
Dim varREsult As Variant

On Error GoTo ErrorHandler
Set objDemand = CreateObject("MSXML2.XMLHTTP")

objDemand.Open "HEAD", strURL, flase
objDemand.Send
varREsult = objDemand.StatusText

Set objDemand = Nothing

If varREsult = "OK" Then
    checkURL = True
End If

ErrorHandler:
End Function

Sub msgMyErrors(nErrors As Integer, Optional noErrorOutput As String = "Script completed with no errors", Optional nameError As String = "Error")
' Help format error

If nErrors = 0 Then
    MsgBox (noErrorOutput)
ElseIf nErrors = 1 Then
    MsgBox ("1 " & nameError & " error detected")
Else
    MsgBox (nErrors & " " & nameError & "s error detected")
End If

End Sub

Function getFileRev(link As String, co As CrossOffice) As String
' Getting a file revesion from doc or excel
' CrossOffice needs to be initialized before using this method

If Right(link, 3) = "xls" Then
    co.setWB (link)
    On Error GoTo Err
    getFileRev = formalRev(CrossOffice.retEXCode("Macro.xlsm!retPropRev"))
    co.closeWB
ElseIf Right(link, 3) = "doc" Then
    On Error GoTo Err
    getFileRev = formalRev(retPropVal(REVISION_PROPERTY_NAME))
Else
Err:
    getFileRev = "Error"
End If

End Function

Function formalRev(rev As String) As String
' Make sure all revs follow the same format B, 00, 01, 02, 03..

If Len(rev) = 1 And rev <> "B" Then
    formalRev = "0" & rev
ElseIf rev = "B" Or rev = "base" Or rev = "Base" Then
    formalRev = "BASE"
End If
    
End Function

Function CutLeftUntil(str As String, cutTo As String, Optional include As Boolean = False, Optional timesToCut As Integer = 1)
' Cuts text from the left until certain String is reached a certain times
' By default the string will be removed
' For example if str = "192.168.1.1" and we cut to "." with include on and 2 times we get "192.168"

For i = 1 To timesToCut
    For counter = Len(str) To 1 Step -1
        If Mid(str, counter, Len(cutTo)) = cutTo Then
            If include Then
                str = Left(str, counter - 1)
            Else
                str = Left(str, counter + Len(cutTo) - 1)
            End If
            GoTo StopCutting
        End If
    Next
StopCutting:
Next

CutLeftUntil = str
End Function

Function CutRightUntil(str As String, cutTo As String, Optional include As Boolean = False, Optional timesToCut As Integer = 1)
' Cuts text from the left until certain String is reached a certain times
' By default the string will be removed
' For example if str = "192.168.1.1" and we cut to "." with include on and 2 times we get "1.1"

For i = 1 To timesToCut
    For counter = 1 To Len(str)
        If Mid(str, counter, Len(cutTo)) = cutTo Then
            If include Then
                str = Right(str, Len(str) - counter - Len(cutTo) + 1)
            Else
                str = Right(str, Len(str) - counter + 1)
            End If
            GoTo StopCutting
        End If
    Next
StopCutting:
Next

CutRightUntil = str
End Function

Sub searchFolder()
' Search all word files in folder to find one which contains certain text

Dim myFile As Variant
Dim conter As Integer
Dim myFolder As String
Dim myMan As Document

myFolder = "C:\User\user\searchfolder"
myFile = Dir(myFolder & "\*.doc")

Do While myFile <> 0

Set myMan = Documents.Open(myFolder & "\" & myFile)
On Error Resume Next
myMan.Unprotect ' Enables serching in headers

If ActiveDocument.Content.Find.Execute("thing to search") Then
    MsgBox myMan.Name
End If

myNan.Close False

myFile = Dir
Loop

End Sub

