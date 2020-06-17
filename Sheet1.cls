VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub CommandButton1_Click()

Dim d As Date
Dim x As Integer

d = CDate(Worksheets("Tasks").Cells(1, 2))

x = DateDiff("d", d, Now)

' MsgBox Now & " <<>> " & d & " -- " & x

' formTaskList
formAndOpenHTML



End Sub

