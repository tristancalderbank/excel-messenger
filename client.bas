Attribute VB_Name = "Module1"
Sub sync_with_server()

    Dim message_window As Range
    Dim root_folder As String
    Dim top_of_messages As String
    Dim bottom_of_messages As String
    
    Dim index As Integer
    Dim update_period As Integer
    
    update_period = 20 ' seconds
    
    top_of_messages = "B3"
    bottom_of_messages = "B23"
    
    root_folder = Application.ActiveWorkbook.Path

    Set message_window = Range(top_of_messages & ":" & bottom_of_messages)

    index = 3
    For Each cell In message_window
        cell.Value = "='" & root_folder & "\[server.xlsm]Sheet1'!B" & index
        index = index + 1
        
        If cell.Value = 0 Then
            cell.Value = ""
        End If
    Next cell


    ActiveWorkbook.Save
    Application.OnTime Now + TimeValue("00:00:" & update_period), "sync_with_server"

End Sub

Private Sub Auto_Open()
    Call sync_with_server
End Sub
