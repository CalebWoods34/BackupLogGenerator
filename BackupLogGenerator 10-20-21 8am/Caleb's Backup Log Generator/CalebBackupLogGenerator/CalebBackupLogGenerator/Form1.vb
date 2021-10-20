Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim UserEnteredPassword As String, CalebPin As String, KevinPin As String, VinhPin As String, MikePin As String
    Dim Foldername As String, strtext As String, strtext2 As String, FileName As String
    Dim xlApp As New Excel.Application, xlWorkBook As Excel.Workbook, xlWorkSheet As Excel.Worksheet, xlformattedapp As New Excel.Application, xlformattedworkbook As Excel.Workbook, xlformattedworksheet As New Excel.Worksheet
    Dim i, j
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click 'This is the label that displays "Backup Log Generator" title
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'This is the command for the button that will allow the user to input their password
        UserEnteredPassword = InputBox("Enter your 4-digit PIN number: ") 'This creates the input box for the user to enter their 4 digit PIN number
        CalebPin = 3189 'These are the pins selected by each authorized user
        KevinPin = 2002
        VinhPin = 1314
        MikePin = 3311
        OpenFileDialog2.InitialDirectory = "M:\Engineering\Engineering staff user files\Caleb"
        OpenFileDialog2.Title = "Select the template you would like to use (.xlsx): "
        OpenFileDialog2.Filter = "Microsoft Excel Worksheet (*.xlsx)|*.xlsx"
        OpenFileDialog1.Title = "Select the backup data log (.csv): " 'This titles the search box 
        OpenFileDialog1.Filter = "Comma Seperated Value File (*.csv)|*.csv"
        OpenFileDialog1.InitialDirectory = "S:\KEVIN\Dougs Automated Cure\AC-CONFIG\Data\BackupLogs\logs" 'This directs the openfiledialog box to search this directory
        If UserEnteredPassword = CalebPin Then 'This is the loop for determining if one of the valid PIN numbers was entered to allow the user access to the backup logs 
            MsgBox("Welcome, Caleb.")
            OpenFileDialog1.ShowDialog()
            OpenFileDialog2.ShowDialog()
            strtext = OpenFileDialog1.FileName
            strtext2 = OpenFileDialog2.FileName
            xlWorkBook = xlApp.Workbooks.Open(strtext) 'This opens the backup log in Excel
            xlformattedworkbook = xlformattedapp.Workbooks.Open(strtext2)
            xlformattedapp.Cells(1, 12) = "PRINT"
            xlformattedapp.Cells(1, 13) = "BY: "
            xlformattedapp.Cells(2, 12) = "Caleb"
            xlformattedapp.Cells(2, 13) = "E405"
            xlApp.Visible = False
            xlformattedapp.Visible = True
            For i = 1 To 1450
                For j = 1 To 50
                    xlformattedapp.Cells(i + 38, j + 1) = xlApp.Cells(i, j) 'this is the loop that prints the values to the excel sheet
                Next
            Next
        ElseIf UserEnteredPassword = KevinPin Then
            MsgBox("Welcome, Kevin.")
            OpenFileDialog1.ShowDialog()
            OpenFileDialog2.ShowDialog()
            strtext = OpenFileDialog1.FileName
            strtext2 = OpenFileDialog2.FileName
            xlWorkBook = xlApp.Workbooks.Open(strtext)
            xlformattedworkbook = xlformattedapp.Workbooks.Open(strtext2)
            xlformattedapp.Cells(1, 12) = "PRINT"
            xlformattedapp.Cells(1, 13) = "BY: "
            xlformattedapp.Cells(2, 12) = "KEVIN"
            xlformattedapp.Cells(2, 13) = "E208"
            xlApp.Visible = False
            xlformattedapp.Visible = True
            For i = 1 To 1450
                For j = 1 To 50
                    xlformattedapp.Cells(i + 38, j + 1) = xlApp.Cells(i, j)
                Next
            Next
        ElseIf UserEnteredPassword = VinhPin Then
            MsgBox("Welcome, Vinh.")
            OpenFileDialog1.ShowDialog()
            OpenFileDialog2.ShowDialog()
            strtext = OpenFileDialog1.FileName
            strtext2 = OpenFileDialog2.FileName
            xlWorkBook = xlApp.Workbooks.Open(strtext)
            xlformattedworkbook = xlformattedapp.Workbooks.Open(strtext2)
            xlformattedapp.Cells(1, 12) = "PRINT"
            xlformattedapp.Cells(1, 13) = "BY: "
            xlformattedapp.Cells(2, 12) = "VINH"
            xlformattedapp.Cells(2, 13) = "E396"
            xlApp.Visible = False
            xlformattedapp.Visible = True
            For i = 1 To 1450
                For j = 1 To 50
                    xlformattedapp.Cells(i + 38, j + 1) = xlApp.Cells(i, j)
                Next
            Next
        ElseIf UserEnteredPassword = MikePin Then
            MsgBox("Welcome, Mike.")
            OpenFileDialog1.ShowDialog()
            OpenFileDialog2.ShowDialog()
            strtext = OpenFileDialog1.FileName
            strtext2 = OpenFileDialog2.FileName
            xlWorkBook = xlApp.Workbooks.Open(strtext)
            xlformattedworkbook = xlformattedapp.Workbooks.Open(strtext2)
            xlformattedapp.Cells(1, 12) = "PRINT"
            xlformattedapp.Cells(1, 13) = "BY: "
            xlformattedapp.Cells(2, 12) = "MIKE"
            xlformattedapp.Cells(2, 13) = "E346"
            xlApp.Visible = False
            xlformattedapp.Visible = True
            For i = 1 To 1450
                For j = 1 To 50
                    xlformattedapp.Cells(i + 38, j + 1) = xlApp.Cells(i, j)
                Next
            Next
        Else
            MsgBox("This is not a valid PIN number.")
        End If
    End Sub

End Class

