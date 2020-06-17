Attribute VB_Name = "docsUtilities"
Option Explicit
Dim listDocs As Collection

Public Sub formDocsList()
    Dim docsRange As Range
    Dim row As Range
    Dim docsObject As docs
    
    Set listDocs = New Collection
    Set docsRange = Worksheets("Docs").Range("B5:F1000")
        
    For Each row In docsRange.rows
        If row.Cells(1) <> "" Then
            Set docsObject = formDocs(row:=row)
            listDocs.Add Item:=docsObject
        End If
    Next row
    
End Sub

Public Function formDocs(row As Range) As docs

        Dim docsSlNo As Long
        Dim docsTitle As String
        Dim docsDescription As String
        Dim docsLink As String
        Dim docsType As String
        
        Dim docsObject As Object
        Set docsObject = New docs
        
        docsSlNo = row.Cells(1)
        docsTitle = row.Cells(2)
        docsDescription = row.Cells(3)
        docsLink = row.Cells(4)
        docsType = row.Cells(5)
        
        With docsObject
                .docsSlNo = docsSlNo
                .docsTitle = docsTitle
                .docsDescription = docsDescription
                .docsLink = docsLink
                .docsType = docsType
        End With

        Set formDocs = docsObject

End Function

Public Function getDocsContent() As String
    Dim doc As docs
    Dim content As String, tempContent As String
    
    Call formDocsList
    For Each doc In listDocs
        tempContent = ""
        tempContent = tempContent & formHTMLColumn(doc.docsSlNo & "). ")
        tempContent = tempContent & formHTMLColumn(doc.docsTitle & "<br/>" & doc.docsDescription)
        tempContent = tempContent & formHTMLColumn("<a href=""" & doc.docsLink & """>" & doc.docsLink & " </a>")
        tempContent = formHTMLRow(tempContent)
        
        content = content & tempContent
    Next doc
    
    getDocsContent = formHTMLTable(content)
    
End Function


