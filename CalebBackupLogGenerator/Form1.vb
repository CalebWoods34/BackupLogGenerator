Public Class Form1
    Dim UserEnteredPassword As String
    Dim pin1 As String
    Dim pin2 As String
    Dim pin3 As String
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click 'This is the label that displays "Backup Log Generator" title

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'This is the command for the button that will allow the user to input their password
        UserEnteredPassword = InputBox("Enter your 4-digit PIN number: ") 'This creates the input box for the user to enter their 4 digit PIN number
        pin1 = 1234
        pin2 = 2345
        pin3 = 3456
        If UserEnteredPassword = pin1 Then
            MsgBox("success")
        ElseIf UserEnteredPassword = pin2 Then
            MsgBox("success")
        ElseIf UserEnteredPassword = pin3 Then
            MsgBox("success")
        Else
            MsgBox("This is not a valid PIN number.")
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
