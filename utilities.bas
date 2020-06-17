Attribute VB_Name = "utilities"
Option Explicit
Dim listHtmlContent As Scripting.Dictionary

Public Sub formAndOpenHTML()

Dim d As Date
Dim x As Integer
Dim sheetsList As Variant
Dim myFile As String, anchorValue As String
Dim htmlContent As String

' d = CDate(Worksheets("Tasks").Cells(1, 2))

' x = DateDiff("d", d, Now)
' sheetsList = listSheets

' MsgBox Now & " <<>> " & d & " -- " & x

htmlContent = formHTMLDoc

        myFile = ActiveWorkbook.Path & "\html\test1.html"
        Open myFile For Output As #1
        Print #1, htmlContent
        Close #1
        
        anchorValue = ActiveWorkbook.Path & "\html\test1.html"
        
        ActiveWorkbook.FollowHyperlink _
        Address:=anchorValue, _
        NewWindow:=True

End Sub

Public Function listSheets() As Variant
    Dim i As Integer
    
    ReDim SNarray(1 To Sheets.Count)
    For i = 1 To Sheets.Count
        SNarray(i) = ThisWorkbook.Sheets(i).Name
    Next
    listSheets = SNarray
End Function

Public Function formHTMLDoc() As String
    Dim htmlFormedContent As String
    Dim htmlTemplate As String
    htmlTemplate = Worksheets("Statics").Cells(5, 17)
    
    Set listHtmlContent = New Scripting.Dictionary
    
    listHtmlContent.Add Key:="Tasks", Item:=getTaskContent
    listHtmlContent.Add Key:="Docs", Item:=getDocsContent
    listHtmlContent.Add Key:="Notes", Item:=getNotesContent
    listHtmlContent.Add Key:="Learn", Item:=getLearnContent
    
    htmlFormedContent = Replace(htmlTemplate, "#BUTTON_NAME#", htmlButtons)
    htmlFormedContent = Replace(htmlFormedContent, "#CONTENT#", htmlContent)
    
    formHTMLDoc = htmlFormedContent
    
End Function

Public Function htmlButtons() As String
    
    Dim htmlVar As Variant
    Dim htmlButtonString As String
    Dim buttonTemplate As String
    
    buttonTemplate = Worksheets("Statics").Cells(3, 17)
    
    For Each htmlVar In listHtmlContent.Keys()
        htmlButtonString = htmlButtonString & Replace(buttonTemplate, "#BUTTON_NAME#", htmlVar)
    Next htmlVar
    
    htmlButtons = htmlButtonString

End Function

Public Function htmlContent() As String
    
    Dim htmlVar As Variant
    Dim htmlButtonString As String
    Dim htmlContentTemplate As String
    Dim tempContent As String
    htmlContentTemplate = Worksheets("Statics").Cells(4, 17)
    
    For Each htmlVar In listHtmlContent.Keys()
        tempContent = htmlContentTemplate
        tempContent = Replace(tempContent, "#BUTTON_NAME#", htmlVar)
        tempContent = Replace(tempContent, "#CONTENT#", listHtmlContent(htmlVar))
        htmlButtonString = htmlButtonString & tempContent
    Next htmlVar
    
    htmlContent = htmlButtonString

End Function

Public Function formHTMLColumn(tempText As String) As String

    tempText = "<td>" & tempText & "</td>"
    formHTMLColumn = tempText

End Function

Public Function formHTMLRow(tempText As String) As String

    tempText = "<tr>" & tempText & "</tr>"
    formHTMLRow = tempText

End Function

Public Function formHTMLTable(tempText As String) As String

    tempText = "<table class=""pure-table"" style=""width:100%"">" & tempText & "</table>"
    formHTMLTable = tempText

End Function

Public Function includeHTMLReplacements(content As String) As String
    Dim formedText As String
    formedText = formMail(content)
    formedText = formPics(formedText)
    formedText = formFiles(formedText)
    
    includeHTMLReplacements = formedText
End Function

Public Function formMail(content As String) As String
    Dim mails() As String
    Dim mailText As String
    
    mailText = SuperMid(content, "<mail>", "</mail>")
    mails = Split(mailText, "||")
    
    If mailText <> "" Then
        mailText = "<mail>" & mailText & "</mail>"
        formMail = Replace(content, mailText, formLinkedHTMLCode(mails, 1))
    Else
        formMail = content
    End If
End Function

Public Function formFiles(content As String) As String
    Dim files() As String
    Dim filesText As String
    
    filesText = SuperMid(content, "<file>", "</file>")
    files = Split(filesText, "||")
    
    If filesText <> "" Then
        filesText = "<file>" & filesText & "</file>"
        formFiles = Replace(content, filesText, formLinkedHTMLCode(files, 2))
    Else
        formFiles = content
    End If
End Function
        
Public Function formPics(content As String) As String
    Dim pics() As String
    Dim picsText As String
    
    picsText = SuperMid(content, "<pic>", "</pic>")
    pics = Split(picsText, "||")
    
    If picsText <> "" Then
        picsText = "<pic>" & picsText & "</pic>"
        formPics = Replace(content, picsText, formLinkedHTMLCode(pics, 3))
    Else
        formPics = content
    End If
End Function

Public Function formLinkedHTMLCode(listObjects() As String, flag As Integer) As String
    Dim linkedText As String
    Dim obj As Variant
    Dim iconText As String
    Dim backgroundColor As String
    Dim linkText As String
    
    If flag = 1 Then
        iconText = "fa fa-envelope-o"
    ElseIf flag = 2 Then
        iconText = "fa fa-folder"
    Else
        iconText = "fa fa-file-photo-o"
    End If
    
    For Each obj In listObjects
        linkText = Split(obj, "link:")(0)
        backgroundColor = ""
        If InStr(linkText, "color:") > 0 Then
            backgroundColor = Split(linkText, "color:")(1)
            linkText = Replace(linkText, "color:" & backgroundColor, "")
        End If
        
        linkedText = linkedText & " <a href=""" & Split(obj, "link:")(1) & """> <button class=""btn""  style=""background-color:" & backgroundColor & " "" ><i class=""" & iconText & """></i> " & linkText & "</button></a> "
    Next obj
    
    formLinkedHTMLCode = "<br/>" & linkedText
End Function
