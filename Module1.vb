Imports System.Data.SqlClient
Imports System.Data

Module Module1
    Public connnectionstring As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\jerem\Patients.mdf;Integrated Security=True"
    Public con As New SqlConnection(connnectionstring)
    Public cmd As New SqlCommand
    Public da As New SqlDataAdapter
    Public dr As SqlDataReader
    Public ds As New DataSet

    Public qry As String = ""


    'search gridview

    Public Function FetchData(ByVal qry As String) As DataSet
        If con.State = 1 Then con.Close()
        con.Open()
        da = New SqlDataAdapter(qry, con)
        ds = New DataSet
        da.Fill(ds)
        Return ds
        con.Close()

    End Function
End Module
