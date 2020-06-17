Attribute VB_Name = "taskUtilities"
Option Explicit
Dim listTasks As Collection

Public Sub formTaskList()
    Dim tasksRange As Range
    Dim row As Range
    Dim taskObject As tasks
    
    Set listTasks = New Collection
    
    Set tasksRange = Worksheets("Tasks").Range("B5:J2000")
        
    For Each row In tasksRange.rows
        If row.Cells(1) <> "" Then
            Set taskObject = formTasks(row:=row)
            listTasks.Add Item:=taskObject
        End If
    Next row
    
    ' MsgBox listTasks.Count
    
    
End Sub

Public Function formTasks(row As Range) As tasks
        
        Dim taskSlNo As Long
        Dim taskAssigned As String, taskType As String, taskTitle As String, taskStatus As String, taskDescription As String
        Dim taskStartTime As Date, taskCompletedTime As Date
        Dim taskDays As Integer
        
        Dim taskObject As Object
        Set taskObject = New tasks
        
        taskSlNo = row.Cells(1)
        taskAssigned = row.Cells(2)
        taskType = row.Cells(3)
        taskTitle = row.Cells(4)
        taskDescription = row.Cells(5)
        taskStatus = row.Cells(6)
        taskStartTime = row.Cells(7)
        taskCompletedTime = row.Cells(8)
        taskDays = row.Cells(9)
         
        With taskObject
                .taskSlNo = taskSlNo
                .taskAssigned = taskAssigned
                .taskType = taskType
                .taskTitle = taskTitle
                .taskDescription = taskDescription
                .taskStatus = taskStatus
                .taskStartTime = taskStartTime
                .taskCompletedTime = taskCompletedTime
                .taskDays = taskDays
        End With

        Set formTasks = taskObject

End Function

Public Function checkIfAlarm(task As tasks) As Boolean

    Dim processDate As Date
    Dim diffDays As Integer
    Dim diffFromToday As Integer
    
    processDate = CDate(Worksheets("Tasks").Cells(1, 2))
    diffDays = DateDiff("d", task.taskStartTime, processDate)
    diffFromToday = DateDiff("d", processDate, Now)
     If task.taskType = "Repititve" Then
        If (diffDays Mod task.taskDays) = 0 Or ((diffDays Mod task.taskDays) + diffFromToday) >= task.taskDays Then
           checkIfAlarm = True
        Else
           checkIfAlarm = False
        End If
    Else
        checkIfAlarm = True
    End If
     
End Function

Public Function getTaskContent() As String
    Dim task As tasks
    Dim content As String, tempContent As String
    Set task = New tasks
    
    Call formTaskList
    For Each task In listTasks
        If (task.taskStatus <> "Completed" And task.taskType <> "Repititve") Or (task.taskStatus <> "Completed" And checkIfAlarm(task)) Then
            tempContent = ""
            tempContent = tempContent & formHTMLColumn(task.taskSlNo & "). ")
            tempContent = tempContent & formHTMLColumn(task.taskTitle & "<br/>" & task.taskDescription)
            tempContent = tempContent & formHTMLColumn("STATUS: " & task.taskStatus & "<br/> " & "ASSIGNED: " & task.taskAssigned)
            tempContent = tempContent & formHTMLColumn(task.taskStartTime)
            tempContent = includeHTMLReplacements(tempContent)
            tempContent = formHTMLRow(tempContent)
            
            content = content & tempContent
        End If
    Next task
    
    getTaskContent = formHTMLTable(content)
End Function

