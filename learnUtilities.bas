Attribute VB_Name = "learnUtilities"
Option Explicit
Dim listLearn As Collection

Public Sub formLearnList()
    Dim learnRange As Range
    Dim row As Range
    Dim learnObject As learn
    
    Set listLearn = New Collection
    Set learnRange = Worksheets("Learn").Range("B4:E1000")
        
    For Each row In learnRange.rows
        If row.Cells(1) <> "" Then
            Set learnObject = formLearn(row:=row)
            listLearn.Add Item:=learnObject
        End If
    Next row
    
End Sub

Public Function formLearn(row As Range) As learn

        Dim learnSlNo As Long
        Dim learnTitle As String
        Dim learnDescription As String
        Dim learnSegment As String
        
        Dim learnObject As Object
        Set learnObject = New learn
        
        learnSlNo = row.Cells(1)
        learnTitle = row.Cells(2)
        learnDescription = row.Cells(3)
        learnSegment = row.Cells(4)
        
        With learnObject
                .learnSlNo = learnSlNo
                .learnTitle = learnTitle
                .learnDescription = learnDescription
                .learnSegment = learnSegment
        End With

        Set formLearn = learnObject

End Function

Public Function getLearnContent() As String
    Dim learn As learn
    Dim content As String, tempContent As String
    
    Call formLearnList
    For Each learn In listLearn
        tempContent = ""
        tempContent = tempContent & formHTMLColumn(learn.learnSlNo & "). ")
        tempContent = tempContent & formHTMLColumn(learn.learnTitle)
        tempContent = tempContent & formHTMLColumn(learn.learnDescription)
        tempContent = formHTMLRow(tempContent)
        
        content = content & tempContent
    Next learn
    
    getLearnContent = formHTMLTable(content)

End Function

