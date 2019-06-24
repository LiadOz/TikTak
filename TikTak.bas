Attribute VB_Name = "TikTak"
' This module is used to automate the check of files containing dozens of links to forms to check for the validity of the file
' The method of excution is:
' 1. Checking if all links work
' 2. Updating the links internal revision to the one written in the file
' 3. Marking expired letters
' 4. If all is working also signing the file

' TableRev class not availiable was used to intuitively put information and iterate tables
' Main table constants
Const SIG_NAME = "admin"
Const NAME_CELL = 1
Const DATE_CELL = 2
Const EXP_DATE_CELL = 3
Const REV_NAME = "מהדורה"
Const SIGN_TABLE = 1
Const LTR_TABLE = 3
Const LTR_TABLE_DATE_CELL = 3
' Form Check table constants
Const REV_TABLE = 2
Const START_ROW = 2
Const ADDRESS_COULMN = 1
Const REV_COLUMN = 5
Const REVESION_PROPERTY_NAME = "מהדורה"

Sub MainTikTak()
' Main call to the TikTak

urlflag = allURLValid
revflag = allFormsOK

If urlflag And revflag Then
    markExpired
    addSigToFile
End If

End Sub

Function FunMainTikTak() As Boolean
' Call to use for intergrating in a more complex automatization like in the MassTikTak module in excel

FunMainTikTak = False

urlflag = allURLValid
revflag = allFormsOK

If urlflag And revflag Then
    markExpired
    addSigToFile
    FunMainTikTak = True
End If
End Function

Function allURLValid() As Boolean
' Checks each link in file, if it doesn't work it is colored red
' Returns True if all links work

allURLValid = True

Dim link As Hyperlink
Dim invalidLinkCount As Integer
invalidLinkCount = 0

Application.ScreenUpdating = False
For Each link In ActiveDocument.Hyperlinks
    link.Range.HighlightColorIndex = wdNoHighlight
    
    If Not completeURLCheck(fixPath(link.address)) Then
        invainvalidLinkCount = invalidLinkCount + 1
        link.Range.HighlightColorIndex = wdRed
    End If
    
Next link
Application.ScreenUpdating = True
If invalidLinkCount <> 0 Then
    allURLValid = False
End If

End Function

Private Function allFormsOK() As Boolean


Application.ScreenUpdating = False

Set co = New CrossOffice
Set formTable = New TableRef
co.Initialize
formTable.Initialize (REV_TABLE)
formTable.setRowNum (START_ROW)
allFormsOK = True

While (formTable.areThereMoreRows)
    formTable.setCellNum (ADDRESS_COULMN)
    newRev = getFileRev(formTable.getLink, co)
    formTable.setCellNum (REV_COLUMN)
    formTable.setString (newRev)
    formTable.nextRow
    
    If newRev = "Error" Then
        allFormsOK = False
    End If
    
Wend

co.destroy
Application.ScreenUpdating = True

End Function

Private Sub markExpired()
' Mark the expired letters

For i = 2 To .Rows.Count
    Set cRow = ActiveDocument.Tables(LTR_TABLE).Rows(i).Cells(LTR_TABLE_DATE_CELL).Range
    cRow.Font.ColorIndex = wdNoHighlight ' Setting the row to not be expired
    If cRow = Null Then GoTo Skip ' If the end was reached
    exDate = createDate(formatTableString(cRow)) ' Getting the expiry date of the letter
    
    If Now() > exDate Then
        cRow.Font.ColorIndex = wdRed ' Setting the row to be expired
    End If
Skip:
Next i
End Sub

Private Function addSigToFile()

Dim signTable As Table
Set signTable = ActiveDocument.Tables(SIGN_TABLE)

emptyRow = findEmptyRow(signTable)
' Creating the new rev date accrding to the last entry
newRev = updateDate(formatTableString(signTable.Rows(emptyRow - 1).Cells(EXP_DATE_CELL).Range))

With signTable.Rows(emptyRow)
    .Cells(NAME_CELL).Range = SIG_NAME
    .Cells(DATE_CELL).Range = updateDate(formatTableString(signTable.Rows(emptyRow - 1).Cells(DATE_CELL).Range))
    .Cells(EXP_DATE_CELL).Range = newRev
End With

Call changeProperties(REV_NAME, newRev)

End Function

Private Function completeURLCheck(myURL As String) As Boolean
' Check full web urls relative url and local urls

completeURLCheck = True

If checkURL(myURL) Then GoTo FunEnd
ElseIf checkURL(ActiveDocument.path & "/" & myURL) Then GoTo FunEnd ' If the url is relative
Else
    On Error Resume Next
    If Dir(myURL, vbDirectory) <> "" Then GoTo FunEnd ' If the url is a local file
    completeURLCheck = False
End If

FunEnd:
End Function

Private Function findEmptyRow(curTable As Table) As String
' Finds the empty newest to to sign in

For i = 1 To curTable.Rows.Count
    If formatTableString(curTable.Rows(i).Cells(NAME_CELL).Range) = "" Then
        findEmptyRow = i
        GoTo ExitFun
    End If
Next i

' If there are no more open spots to sign add another one
curTable.Rows.Add
findEmptyRow = curTable.Rows.Count

ExitFun:
End Function

Private Function updateDate(curDate As String) As String
' Returns the date after update

On Error Resume Next
dateYear = CInt(CutRightUntil(curDate, "/", True, 2)) + 1
updateDate = CutLeftUntil(curDate, "/", False, 1) & dateYear

End Function


Private Function createDate(myDate As String) As Date
' Creates a date from string

dd = CutLeftUntil(myDate, "/", True, 2)
mm = CutRightUntil(CutLeftUntil(myDate, "/", True, 1), "/", True)
yyyy = CutRightUntil(myDate, "/", True, 2)
createDate = Now()

If Len(dd) < 2 Then
    dd = "0" & dd
End If
If Len(mm) < 2 Then
    mm = "0" & mm
End If
If Len(yyyy) < 2 Then
    yyyy = "00" & yyyy
End If

On Error Resume Next
createDate = CDate(dd & "/" & mm & "/" & yyyy)

End Function

Private Function fixPath(fixstr As String)
' Converts relative path of string to a full link

path = ActiveDocument.path

back = 0
For i = 1 To Len(fixstr)
    On Error Resume Next
    If Mid(fixstr, counter, 2) = ".." Then
        back = back + 1
    End If
    If Mid(fixstr, counter, 2) <> "./" Then GoTo CountFinish
Next
CountFinish:

For i = back To 1 Step -1
    For j = Len(path) To 1 Step -1
        If Mid(path, counter, 1) = "/" Then
            path = Left(path, counter - 1)
            GoTo StopCutting
        End If
    Next
StopCutting:
Next

If back = 0 Then
    fixPath -fixstr
Else
    fixPath = docPath & "/" & Right(fixstr, Len(fixstr) - back * 3)
End If

End Function
