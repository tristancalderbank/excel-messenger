Attribute VB_Name = "Module1"
Function write_message(message)
    
    Dim current_messages As range
    Set current_messages = range("B3:B23")
    
    For Each cell In current_messages
        If cell.Value = message Then
            write_message = False
            Exit Function
        End If
    Next cell
    
    If message = "" Or message = 0 Then
        write_message = False
        Exit Function
    
    End If
    
    ' shift messages up
    Dim shifted_up As range
    Dim old_messages As range
    Dim new_message_cell As range
    
    
    Set shifted_up = range("B3:B22")
    Set old_messages = range("B4:B23")
    
    shifted_up = old_messages.Value
    
    Set new_message_cell = range("B23")
    
    new_message_cell.Value = message
    
    write_message = True

End Function


Sub run_server()
   
    Dim buffer_cell As range
    Dim message As String
    Dim nickname As String
    Dim root_folder As String
    Dim number_of_clients As Integer
    Dim update_period As Integer
    
    update_period = 20
    number_of_clients = 1
    root_folder = Application.ActiveWorkbook.Path
      
    Set buffer_cell = range("B25")

    For i = 0 To number_of_clients - 1
        buffer_cell.Value = "='" & root_folder & "\[client_" & i & ".xlsm]Sheet1'!$B$25"
        message = buffer_cell.Value
        buffer_cell.Value = "='" & root_folder & "\[client_" & i & ".xlsm]Sheet1'!$D$5"
        nickname = buffer_cell.Value
        write_message (nickname & ": " & message)
    Next i

    ActiveWorkbook.Save
    Application.OnTime Now + TimeValue("00:00:" & update_period), "run_server"
    
End Sub

Private Sub Auto_Open()
    Call run_server
End Sub
