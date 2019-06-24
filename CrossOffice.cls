VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Module to perform actions across multiple office applications, needed Office Excel and Word Object Libarary reference

Private ex As Excel.Application
Private wd As Word.Application
Private wb As Workbook
Private doc As Document
Private macroKey As Object

Private wdPath As String
Private exPath As String

Sub initiallize()

Set ex = Excel.Application
Set wd = Word.Application

ex.Visible = False
wd.Visible = False

End Sub

Sub setWDPath(path As String)
wdPath = path
End Sub

Sub setExPath(path As String)
exPath = path
End Sub

Sub setDoc(path As String)
Set doc = wd.Documents.Open(path)
End Sub

Sub setWB(path As String)
Set wb = ex.Workbooks.Open(path)
End Sub

Sub closeDoc()
doc.Close (False)
End Sub

Sub closeWB()
wb.Close (False)
End Sub

Sub executeWDCode(macro As String)
On Error Resume Next
wd.Run (macro)
End Sub

Function retWDCode(macro As String)
' To return value from macro
On Error Resume Next
retWDCode = wd.Run(macro)
End Function

Sub executeEXCode(macro As String)
On Error Resume Next
ex.Run (macro)
End Sub

Function retEXCode(macro As String)
' To return value from macro
On Error Resume Next
retWDCode = ex.Run(macro)
End Function

Function destroy()
wd.Quit
ex.Quit
End Function

