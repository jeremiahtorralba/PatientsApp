Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices
Imports Microsoft.Identity

Public Class Form1

    'Inherits System.Windows.Forms.Form
    'Create ADO.NET objects.'
    Private defaultConn As SqlConnection
    Private myCmd As SqlCommand
    Private myCmd2 As SqlCommand
    Private myReader As SqlDataReader
    Private myReader2 As SqlDataReader
    Private results As String

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        defaultConn = New SqlConnection("Initial Catalog=Patients;" & "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\jerem\Patients.mdf;Integrated Security=True;")

        myCmd = defaultConn.CreateCommand 'Create a Connection object.'
        myCmd.CommandText = "SELECT Id, Assigned_ID, Outbreak Associated, Age Group, Neighbourhood Name, FSA, Source of Infection, Classification, Episode Date, Reported Date, Client Gender, Outcome, Currently Hospitalized, Currently in ICU, Currently Intubated, Ever Hospitalized, Ever in ICU, Ever Intubated, Year, Month FROM Patients"
        defaultConn.Open() 'Open the connection.'
        myReader = myCmd.ExecuteReader()

        'Close the reader and the database connection.'
        myReader.Close()
        defaultConn.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'PatientsDataSet.Cases' table. You can move, or remove it, as needed.
        Me.CasesTableAdapter.Fill(Me.PatientsDataSet.Cases)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False

        'LoadDatainGrid()

    End Sub

    'Public Sub LoadDatainGrid()
    'If con.State = 1 Then con.Close()
    '   ds.Clear()
    '  DataGridView1.DataSource = Nothing
    ' qry = "Select * from Cases with (nolock)"
    'ds = FetchData(qry)
    'If ds.Tables(0).Rows.Count > 0 Then
    '       DataGridView1.DataSource = ds.Tables(0)
    'Else
    '       MessageBox.Show("Data not found....")
    'End If
    'End Sub

    'Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    '        defaultConn = New SqlConnection("Database=Patients;" & "Initial Catalog=Patients;" & "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\jerem\Patients.mdf;Integrated Security=True;")

    '    Try
    '           defaultConn.Open()
    '  Dim cmd As New SqlCommand("SELECT * FROM Cases WHERE Id like '%" & TextBox4.Text & "%' OR Id like '%" & TextBox4.Text & "%'", defaultConn)
    '         myReader2 = cmd.ExecuteReader
    'While myReader2.Read
    '           DataGridView2.Rows.Add(myReader2.Item("ID"), myReader2.Item("Assigned_ID"), myReader2.Item("Outbreak Associated"), myReader2.Item("Age Group"), myReader2.Item("Neighbourhood Name"), myReader2.Item("FSA"), myReader2.Item("Source Of Infection"), myReader2.Item("Classification"), myReader2.Item("Episode Date"), myReader2.Item("Reported Date"), myReader2.Item("Client Gender"), myReader2.Item("Outcome"), myReader2.Item("Currently Hospitalized"), myReader2.Item("Currently In ICU"), myReader2.Item("Currently Intubated"), myReader2.Item("Ever Hospitalized"), myReader2.Item("Ever In ICU"), myReader2.Item("Ever Intubated"), myReader2.Item("Year"), myReader2.Item("Month"))
    'End While
    '       myReader2.Dispose()
    'Catch ex As Exception
    '       MsgBox(ex.Message)
    'End Try
    '   defaultConn.Close()
    'End Sub
End Class