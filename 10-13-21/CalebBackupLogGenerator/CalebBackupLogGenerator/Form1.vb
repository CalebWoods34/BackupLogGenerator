Public Class Form1
    Dim UserEnteredPassword As String 'This is the declaration section for the variables that this program uses
    Dim CalebPin As String
    Dim KevinPin As String
    Dim VinhPin As String
    Dim MikePin As String
    Dim Foldername As String
    Dim strtext As String
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click 'This is the label that displays "Backup Log Generator" title
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'This is the command for the button that will allow the user to input their password
        UserEnteredPassword = InputBox("Enter your 4-digit PIN number: ") 'This creates the input box for the user to enter their 4 digit PIN number
        CalebPin = 3189 'These are the pins selected by each authorized user
        KevinPin = 2002
        VinhPin = 1314
        MikePin = 3311
        If UserEnteredPassword = CalebPin Then 'This is the loop for determining if one of the valid PIN numbers was entered to allow the user access to the backup logs 
            MsgBox("Welcome, Caleb. ")
            OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs" 'This directs the openfiledialog box to search this directory
            OpenFileDialog1.Title = "Select the backup log to open: " 'This titles the search box 
            OpenFileDialog1.ShowDialog()
            strtext = OpenFileDialog1.FileName
        ElseIf UserEnteredPassword = KevinPin Then
            MsgBox("Welcome, Kevin.")
            OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs"
            OpenFileDialog1.Title = "Select the backup log to open: "
            OpenFileDialog1.ShowDialog()
            strtext = OpenFileDialog1.FileName
        ElseIf UserEnteredPassword = VinhPin Then
            MsgBox("Welcome, Vinh.")
            OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs"
            OpenFileDialog1.Title = "Select the backup log to open: "
            OpenFileDialog1.ShowDialog()
            strtext = OpenFileDialog1.FileName
        ElseIf UserEnteredPassword = MikePin Then
            MsgBox("Welcome, Mike.")
            OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs"
            OpenFileDialog1.Title = "Select the backup log to open: "
            OpenFileDialog1.ShowDialog()
            strtext = OpenFileDialog1.FileName
        Else
            MsgBox("This is not a valid PIN number.")
        End If
    End Sub

End Class
