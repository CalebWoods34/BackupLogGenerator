Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim UserEnteredPassword As String 'This is the declaration section for the variables that this program uses
    Dim CalebPin As String
    Dim KevinPin As String
    Dim VinhPin As String
    Dim MikePin As String
    Dim Foldername As String
    Dim strtext As String
    Dim FileName As String
    Dim xlApp As New Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim xlFinalLogDraftApp As New Excel.Application
    Dim xlFinalLogDraftWorkBook As Excel.Workbook
    Dim xlFinalLogDraftWorkSheet As Excel.Worksheet
    Dim xlFinalLogDraftRange As Excel.Range
    Dim i
    Dim j
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click 'This is the label that displays "Backup Log Generator" title
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'This is the command for the button that will allow the user to input their password
        UserEnteredPassword = InputBox("Enter your 4-digit PIN number: ") 'This creates the input box for the user to enter their 4 digit PIN number
        CalebPin = 3189 'These are the pins selected by each authorized user
        KevinPin = 2002
        VinhPin = 1314
        MikePin = 3311
        xlFinalLogDraftWorkBook = xlFinalLogDraftApp.Workbooks.Add
        xlFinalLogDraftWorkSheet = xlFinalLogDraftWorkBook.ActiveSheet
        xlFinalLogDraftApp.Cells(1, 1) = "BACKUP"
        xlFinalLogDraftApp.Cells(1, 2) = "LOG"
        xlFinalLogDraftApp.Cells(1, 3) = Nothing
        xlFinalLogDraftApp.Cells(1, 4) = "PRINT BY: "
        If UserEnteredPassword = CalebPin Then 'This is the loop for determining if one of the valid PIN numbers was entered to allow the user access to the backup logs 
            MsgBox("Welcome, Caleb.")
            OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs" 'This directs the openfiledialog box to search this directory
            OpenFileDialog1.Title = "Select the backup data log (.csv): " 'This titles the search box 
            OpenFileDialog1.Filter = "Comma Seperated Value File (*.csv)|*.csv"
            OpenFileDialog1.ShowDialog()
            strtext = OpenFileDialog1.FileName
            xlWorkBook = xlApp.Workbooks.Open(strtext) 'This opens the backup log in Excel
            xlApp.Visible = True 'This makes the backup log visible, SUBJECT TO CHANGE in order to remove security loophole
            xlFinalLogDraftApp.Cells(1, 5) = "CALEB W."
            xlFinalLogDraftApp.Visible = True 'This makes the reformatted log visible, SUBJECT TO CHANGE if there is a way to automatically print to a PDF file for security purposes
            For i = 1 To 1441
                For j = 1 To 50
                    xlFinalLogDraftApp.Cells(i, j) = xlApp.Cells(i, j)
                Next
            Next
        ElseIf UserEnteredPassword = KevinPin Then
            MsgBox("Welcome, Kevin.")
            OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs"
            OpenFileDialog1.Title = "Select the backup data log (.csv): "
            OpenFileDialog1.Filter = "Comma Seperated Value File (*.csv)|*.csv"
            OpenFileDialog1.ShowDialog()
            strtext = OpenFileDialog1.FileName
            xlWorkBook = xlApp.Workbooks.Open(strtext)
            xlApp.Visible = True
            xlFinalLogDraftApp.Cells(1, 5) = "KEVIN F."
            xlFinalLogDraftApp.Visible = True
        ElseIf UserEnteredPassword = VinhPin Then
            MsgBox("Welcome, Vinh.")
            OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs"
            OpenFileDialog1.Title = "Select the backup data log (.csv): "
            OpenFileDialog1.Filter = "Comma Seperated Value File (*.csv)|*.csv"
            OpenFileDialog1.ShowDialog()
            strtext = OpenFileDialog1.FileName
            xlWorkBook = xlApp.Workbooks.Open(strtext)
            xlApp.Visible = True
            xlFinalLogDraftApp.Cells(1, 5) = "VINH N."
            xlFinalLogDraftApp.Visible = True
        ElseIf UserEnteredPassword = MikePin Then
            MsgBox("Welcome, Mike.")
            OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs"
            OpenFileDialog1.Title = "Select the backup data log (.csv):  "
            OpenFileDialog1.Filter = "Comma Seperated Value File (*.csv)|*.csv"
            OpenFileDialog1.ShowDialog()
            strtext = OpenFileDialog1.FileName
            xlWorkBook = xlApp.Workbooks.Open(strtext)
            xlApp.Visible = True
            xlFinalLogDraftApp.Cells(1, 5) = "MIKE D."
            xlFinalLogDraftApp.Visible = True
        Else
            MsgBox("This is not a valid PIN number.")
        End If
    End Sub
End Class
