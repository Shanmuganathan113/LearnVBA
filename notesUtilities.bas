Attribute VB_Name = "notesUtilities"
Option Explicit
Dim listNotes As Collection

Public Sub formNotesList()
    Dim notesRange As Range
    Dim row As Range
    Dim notesObject As Notes
    
    Set listNotes = New Collection
    Set notesRange = Worksheets("Notes").Range("B4:E1000")
        
    For Each row In notesRange.rows
        If row.Cells(1) <> "" Then
            Set notesObject = formNotes(row:=row)
            listNotes.Add Item:=notesObject
        End If
    Next row
    
End Sub

Public Function formNotes(row As Range) As Notes

        Dim notesSlNo As Long
        Dim notesTitle As String
        Dim notesDescription As String
        Dim notesSegment As String
        
        Dim notesObject As Object
        Set notesObject = New Notes
        
        notesSlNo = row.Cells(1)
        notesTitle = row.Cells(2)
        notesDescription = row.Cells(3)
        notesSegment = row.Cells(4)
        
        With notesObject
                .notesSlNo = notesSlNo
                .notesTitle = notesTitle
                .notesDescription = notesDescription
                .notesSegment = notesSegment
        End With

        Set formNotes = notesObject

End Function

Public Function getNotesContent() As String
    Dim note As Notes
    Dim content As String, tempContent As String
    
    Call formNotesList
    For Each note In listNotes
        tempContent = ""
        tempContent = tempContent & formHTMLColumn(note.notesSlNo & "). ")
        tempContent = tempContent & formHTMLColumn(note.notesTitle)
        tempContent = tempContent & formHTMLColumn(note.notesDescription)
        tempContent = formHTMLRow(tempContent)
        
        content = content & tempContent
    Next note
    
    getNotesContent = formHTMLTable(content)

End Function
