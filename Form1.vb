Imports Microsoft.Data.SqlClient

Public Class Form1

    Dim conn = New SqlConnection("Initial Catalog=Patients;" & "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\jerem\Patients.mdf;Integrated Security=True;")
    Dim i As Integer
    Dim dr As SqlDataReader
    Dim nt = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView_Load()
        Clear()
    End Sub

    Public Sub DataGridView_Load()
        DataGridView1.Rows.Clear()
        Try
            conn.Open()
            Dim cmd As New SqlCommand("SELECT * FROM Cases", conn)
            dr = cmd.ExecuteReader
            While dr.Read
                DataGridView1.Rows.Add(dr.Item("Id"), dr.Item("Assigned_ID"), dr.Item("Outbreak_Associated"), dr.Item("Age_Group"), dr.Item("Neighbourhood_Name"),
                    dr.Item("FSA"), dr.Item("Source_of_Infection"), dr.Item("Classification"), dr.Item("Episode_Date"), dr.Item("Reported_Date"),
                    dr.Item("Client_Gender"), dr.Item("Outcome"), dr.Item("Currently_Hospitalized"), dr.Item("Currently_in_ICU"), dr.Item("Currently_Intubated"),
                    dr.Item("Ever_Hospitalized"), dr.Item("Ever_in_ICU"), dr.Item("Ever_Intubated"), dr.Item("Year"), dr.Item("Month"))
            End While
            dr.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub AddButton_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Add()
    End Sub

    'SET IDENTITY_INSERT Cases ON; 
    'Id
    '(SELECT ISNULL(MAX(@Id)+1,1) FROM Cases)

    '(SELECT MAX(@Id)+1)
    Public Sub Add()
        Dim add As String = "SET IDENTITY_INSERT Cases ON; INSERT INTO Cases (Id, Assigned_ID, Outbreak_Associated, Age_Group, Neighbourhood_Name, FSA, Source_of_Infection, 
            Classification, Episode_Date, Reported_Date, Client_Gender, Outcome, Currently_Hospitalized, Currently_in_ICU, Currently_Intubated,              
            Ever_Hospitalized, Ever_in_ICU, Ever_Intubated, Year, Month) VALUES((SELECT MAX(@Id)), @Assigned_ID, @Outbreak_Associated, @Age_Group, @Neighbourhood_Name, @FSA, @Source_of_Infection, @Classification, @Episode_Date, 
            @Reported_Date, @Client_Gender, @Outcome, @Currently_Hospitalized, @Currently_in_ICU, @Currently_Intubated, @Ever_Hospitalized, @Ever_In_ICU, 
            @Ever_Intubated, @Year, @Month)"
        Try
            Using cmd As New SqlCommand(add, conn)
                conn.Open()
                cmd.Parameters.Clear()
                'cmd.Parameters.AddWithValue("@Id", SqlDbType.Int)
                cmd.Parameters.AddWithValue("@Id", TextBox2.Text.ToString())
                cmd.Parameters.AddWithValue("@Assigned_ID", TextBox2.Text.ToString())
                cmd.Parameters.AddWithValue("@Outbreak_Associated", MaskedTextBox1.Text.ToString())
                cmd.Parameters.AddWithValue("@Age_Group", MaskedTextBox2.Text.ToString())
                cmd.Parameters.AddWithValue("@Neighbourhood_Name", MaskedTextBox3.Text.ToString())
                cmd.Parameters.AddWithValue("@FSA", MaskedTextBox4.Text.ToString())
                cmd.Parameters.AddWithValue("@Source_of_Infection", MaskedTextBox5.Text.ToString())
                cmd.Parameters.AddWithValue("@Classification", MaskedTextBox6.Text.ToString())
                cmd.Parameters.AddWithValue("@Episode_Date", MaskedTextBox7.Text.ToString())
                cmd.Parameters.AddWithValue("@Reported_Date", MaskedTextBox8.Text.ToString())
                cmd.Parameters.AddWithValue("@Client_Gender", MaskedTextBox9.Text.ToString())
                cmd.Parameters.AddWithValue("@Outcome", MaskedTextBox10.Text.ToString())
                cmd.Parameters.AddWithValue("@Currently_Hospitalized", MaskedTextBox11.Text.ToString())
                cmd.Parameters.AddWithValue("@Currently_in_ICU", MaskedTextBox12.Text.ToString())
                cmd.Parameters.AddWithValue("@Currently_Intubated", MaskedTextBox13.Text.ToString())
                cmd.Parameters.AddWithValue("@Ever_Hospitalized", MaskedTextBox14.Text.ToString())
                cmd.Parameters.AddWithValue("@Ever_in_ICU", MaskedTextBox15.Text.ToString())
                cmd.Parameters.AddWithValue("@Ever_Intubated", MaskedTextBox16.Text.ToString())
                cmd.Parameters.AddWithValue("@Year", MaskedTextBox17.Text.ToString())
                cmd.Parameters.AddWithValue("@Month", MaskedTextBox18.Text.ToString())
                'End While
                'SET IDENTITY_INSERT to ON
                'SET IDENTITY_INSERT Cases ON;
                'Go

                i = cmd.ExecuteNonQuery
                If i > 0 Then
                    MessageBox.Show("New Case Saved!", "Add Patient Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("New Case Failed!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                End If
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
        Clear()
        DataGridView_Load()
    End Sub

    'Clear Textboxes
    Public Sub Clear()
        TextBox2.Clear()
        MaskedTextBox1.Clear()
        MaskedTextBox2.Clear()
        MaskedTextBox3.Clear()
        MaskedTextBox4.Clear()
        MaskedTextBox5.Clear()
        MaskedTextBox6.Clear()
        MaskedTextBox7.Clear()
        MaskedTextBox8.Clear()
        MaskedTextBox9.Clear()
        MaskedTextBox10.Clear()
        MaskedTextBox11.Clear()
        MaskedTextBox12.Clear()
        MaskedTextBox13.Clear()
        MaskedTextBox14.Clear()
        MaskedTextBox15.Clear()
        MaskedTextBox16.Clear()
        MaskedTextBox17.Clear()
        MaskedTextBox18.Clear()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        i = DataGridView1.CurrentRow.Cells(0).Value
        TextBox2.Text = DataGridView1.CurrentRow.Cells(1).Value
        MaskedTextBox1.Text = DataGridView1.CurrentRow.Cells(2).Value
        MaskedTextBox2.Text = DataGridView1.CurrentRow.Cells(3).Value
        MaskedTextBox3.Text = DataGridView1.CurrentRow.Cells(4).Value
        MaskedTextBox4.Text = DataGridView1.CurrentRow.Cells(5).Value
        MaskedTextBox5.Text = DataGridView1.CurrentRow.Cells(6).Value
        MaskedTextBox6.Text = DataGridView1.CurrentRow.Cells(7).Value
        MaskedTextBox7.Text = DataGridView1.CurrentRow.Cells(8).Value
        MaskedTextBox8.Text = DataGridView1.CurrentRow.Cells(9).Value
        MaskedTextBox9.Text = DataGridView1.CurrentRow.Cells(10).Value
        MaskedTextBox10.Text = DataGridView1.CurrentRow.Cells(11).Value
        MaskedTextBox11.Text = DataGridView1.CurrentRow.Cells(12).Value
        MaskedTextBox12.Text = DataGridView1.CurrentRow.Cells(13).Value
        MaskedTextBox13.Text = DataGridView1.CurrentRow.Cells(14).Value
        MaskedTextBox14.Text = DataGridView1.CurrentRow.Cells(15).Value
        MaskedTextBox15.Text = DataGridView1.CurrentRow.Cells(16).Value
        MaskedTextBox16.Text = DataGridView1.CurrentRow.Cells(17).Value
        MaskedTextBox17.Text = DataGridView1.CurrentRow.Cells(18).Value
        MaskedTextBox18.Text = DataGridView1.CurrentRow.Cells(19).Value

        'datepicker.text
        'checkbox.Checked

        TextBox2.ReadOnly = False
        Button1.Enabled = False

    End Sub

    Sub Edit()
        Dim edit As String = "UPDATE Cases SET Assigned_ID=@Assigned_ID, Outbreak_Associated=@Outbreak_Associated, Age_Group=@Age_Group, Neighbourhood_Name=@Neighbourhood_Name, FSA=@FSA, Source_of_Infection=@Source_of_Infection, Classification=@Classification, Episode_Date=@Episode_Date, Reported_Date=@Reported_Date, Client_Gender=@Client_Gender, Outcome=@Outcome, Currently_Hospitalized=@Currently_Hospitalized, Currently_in_ICU=@Currently_in_ICU, Currently_Intubated=@Currently_Intubated, Ever_Hospitalized=@Ever_Hospitalized, Ever_in_ICU=@Ever_in_ICU, Ever_Intubated=@Ever_Intubated, Year=@Year, Month=@Month WHERE Assigned_ID=@Assigned_ID"
        Try
            Using cmd As New SqlCommand(edit, conn)
                conn.Open()
                cmd.Parameters.Clear()
                'cmd.Parameters.AddWithValue("@Id", "Id")
                cmd.Parameters.AddWithValue("@Assigned_ID", TextBox2.Text.ToString())
                cmd.Parameters.AddWithValue("@Outbreak_Associated", MaskedTextBox1.Text.ToString())
                cmd.Parameters.AddWithValue("@Age_Group", MaskedTextBox2.Text.ToString())
                cmd.Parameters.AddWithValue("@Neighbourhood_Name", MaskedTextBox3.Text.ToString())
                cmd.Parameters.AddWithValue("@FSA", MaskedTextBox4.Text.ToString())
                cmd.Parameters.AddWithValue("@Source_of_Infection", MaskedTextBox5.Text.ToString())
                cmd.Parameters.AddWithValue("@Classification", MaskedTextBox6.Text.ToString())
                cmd.Parameters.AddWithValue("@Episode_Date", MaskedTextBox7.Text.ToString())
                cmd.Parameters.AddWithValue("@Reported_Date", MaskedTextBox8.Text.ToString())
                cmd.Parameters.AddWithValue("@Client_Gender", MaskedTextBox9.Text.ToString())
                cmd.Parameters.AddWithValue("@Outcome", MaskedTextBox10.Text.ToString())
                cmd.Parameters.AddWithValue("@Currently_Hospitalized", MaskedTextBox11.Text.ToString())
                cmd.Parameters.AddWithValue("@Currently_in_ICU", MaskedTextBox12.Text.ToString())
                cmd.Parameters.AddWithValue("@Currently_Intubated", MaskedTextBox13.Text.ToString())
                cmd.Parameters.AddWithValue("@Ever_Hospitalized", MaskedTextBox14.Text.ToString())
                cmd.Parameters.AddWithValue("@Ever_in_ICU", MaskedTextBox15.Text.ToString())
                cmd.Parameters.AddWithValue("@Ever_Intubated", MaskedTextBox16.Text.ToString())
                cmd.Parameters.AddWithValue("@Year", MaskedTextBox17.Text.ToString())
                cmd.Parameters.AddWithValue("@Month", MaskedTextBox18.Text.ToString())
                i = cmd.ExecuteNonQuery
                If i > 0 Then
                    MessageBox.Show("Case Updated!", "Update Patient Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Case Update Failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
        Clear()
        DataGridView_Load()
    End Sub

    Private Sub EditButton_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Edit()
    End Sub

    ' Query to Increment Id
    'ALTER TABLE Cases
    'ADD Id int IDENTITY(1,1) PRIMARY KEY NOT NULL;

    Public Sub Delete()
        TextBox2.ReadOnly = False
        DataGridView1.MultiSelect = True
        If MsgBox("Do You Want To Delete This Case?", MsgBoxStyle.Question + vbYesNo) = vbYes Then
            Dim delete As String = "DELETE FROM Cases WHERE Id='" & TextBox2.Text & "'"
            Try
                Using cmd As New SqlCommand(delete, conn)
                    conn.Open()
                    cmd.Parameters.Clear()
                    cmd.Parameters.AddWithValue("@Id", DataGridView1.SelectedRows(0).Cells(1).Value)
                    i = cmd.ExecuteNonQuery
                    If i > 0 Then
                        MessageBox.Show("Case Deleted!", "Delete Patient Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Case Deletion Failed, Please Select Unique Case!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    End If
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                conn.Close()
            End Try
            Clear()
            DataGridView_Load()
        Else
            Return
        End If
    End Sub

    Private Sub DeleteButton_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox2.ReadOnly = False
        Delete()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Clear()
        TextBox2.ReadOnly = False
        Button1.Enabled = True
    End Sub

    Private Sub SearchBox_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        Clear()
    End Sub

    Private Sub SearchBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        DataGridView1.Rows.Clear()
        Try
            conn.Open()
            Dim cmd As New SqlCommand("SELECT * FROM Cases WHERE Id LIKE '%" & TextBox3.Text & "%' OR Assigned_ID LIKE '%" & TextBox3.Text & "%'
                OR Outbreak_Associated LIKE '%" & TextBox3.Text & "%' OR Age_Group LIKE '%" & TextBox3.Text & "%' OR Neighbourhood_Name LIKE '%" & TextBox3.Text & "%' 
                OR FSA LIKE '%" & TextBox3.Text & "%' OR Source_of_Infection LIKE '%" & TextBox3.Text & "%' OR Classification LIKE '%" & TextBox3.Text & "%' 
                OR Episode_Date LIKE '%" & TextBox3.Text & "%' OR Reported_Date LIKE '%" & TextBox3.Text & "%' OR Client_Gender LIKE '%" & TextBox3.Text & "%'
                OR Outcome LIKE '%" & TextBox3.Text & "%' OR Currently_Hospitalized LIKE '%" & TextBox3.Text & "%' OR Currently_in_ICU LIKE '%" & TextBox3.Text & "%' 
                OR Currently_Intubated LIKE '%" & TextBox3.Text & "%' OR Ever_Hospitalized LIKE '%" & TextBox3.Text & "%' OR Ever_in_ICU LIKE '%" & TextBox3.Text & "%'
                OR Ever_Intubated LIKE '%" & TextBox3.Text & "%' OR Year LIKE '%" & TextBox3.Text & "%' OR Month LIKE '%" & TextBox3.Text & "%'", conn)
            dr = cmd.ExecuteReader
            While dr.Read
                DataGridView1.Rows.Add(dr.Item("Id"), dr.Item("Assigned_ID"), dr.Item("Outbreak_Associated"), dr.Item("Age_Group"), dr.Item("Neighbourhood_Name"),
                    dr.Item("FSA"), dr.Item("Source_of_Infection"), dr.Item("Classification"), dr.Item("Episode_Date"), dr.Item("Reported_Date"),
                    dr.Item("Client_Gender"), dr.Item("Outcome"), dr.Item("Currently_Hospitalized"), dr.Item("Currently_in_ICU"), dr.Item("Currently_Intubated"),
                    dr.Item("Ever_Hospitalized"), dr.Item("Ever_in_ICU"), dr.Item("Ever_Intubated"), dr.Item("Year"), dr.Item("Month"))
            End While
            dr.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click
        Button1.Enabled = True
    End Sub
End Class