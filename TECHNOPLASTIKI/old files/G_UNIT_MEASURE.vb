Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class gergates


    'Create connection
    Dim conn As OleDbConnection

    'create data adapter
    Dim da As OleDbDataAdapter ' SqlDataAdapter

    'create dataset
    Dim ds As DataSet = New DataSet

    'Set up connection string
    Dim cnString As String
    'Public Property PLATOS() As Integer
    '    Get
    '        Return VALUE
    '    End Get
    '    Set(ByVal value As Integer)

    '    End Set
    'End Property

    Private STHLH_ID As Integer = 1
    Public Property STHLHTOY_ID() As Integer
        Get
            Return STHLH_ID
        End Get
        Set(ByVal Value As Integer)
            If Value < 1 Or Value > 12 Then
                ' Error processing for invalid value. 
            Else
                STHLH_ID = Value
            End If
        End Set
    End Property





    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' FormAdd("FrmUNIT_MEASURE_ADD")
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            da.Update(ds, "PEL")
        Catch ex As Exception
            MsgBox("δεν αποθηκευτηκε" + ex.Message)
        End Try

        Me.Close()
    End Sub

    Private Sub FrmUNIT_MEASURE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '  sqlSTR = "SELECT * FROM TBL_Unit_Measure"
        ' FillListView(ExecuteSQLQuery(sqlSTR), lstunit, 0)
        paint_ergasies()
    End Sub
    Sub paint_ergasies()


        'Imports System.Data.OleDb
        'Imports System.Data.SqlClient

        'Public Class ergates


        'Create connection
        '       Dim conn As SqlConnection

        'create data adapter
        '      Dim da As SqlDataAdapter

        'create dataset
        '     Dim ds As DataSet = New DataSet

        'Set up connection string
        '    Dim cnString As String


        'gConnect = "Provider=SQLOLEDB;" & _
        '                      "Data Source=" & Split(tmpStr, ":")(0) & _
        '                      ";Network=" & Split(tmpStr, ":")(1) & _
        '                      ";Server=" & Split(tmpStr, ":")(1) & _
        '                      ";Initial Catalog=thermo" & _
        '                      ";User Id=" & Split(tmpStr, ":")(2) & _
        '                      ";Password=" & Split(tmpStr, ":")(3)
        'Dim sqlCon As New OleDbConnection
        'sqlCon.ConnectionString = gConnect
        'sqlCon.Open()

        cnString = gConnect ' "Data Source=localhost\SQLEXPRESS;Integrated Security=True;database=thermo"
        'Str_Connection = cnString
        Dim SQLqry
        SQLqry = Label1.Text '"SELECT NAME,N1,ID FROM ERGATES " ' ORDER BY HME "
        'conn = New SqlConnection(cnString)

        Dim conn As New OleDbConnection
        conn.ConnectionString = gConnect
        conn.Open()



        Try
            ' Open connection
            'conn.Open()

            da = New OleDbDataAdapter(SQLqry, conn)

            'create command builder
            Dim cb As OleDbCommandBuilder = New OleDbCommandBuilder(da)
            ds.Clear()
            'fill dataset
            da.Fill(ds, "PEL")
            GridView1.ClearSelection()
            GridView1.DataSource = ds
            GridView1.DataMember = "PEL"

            'GridView1.Columns(STHLHTOY_ID).Width = 0
            GridView1.Columns(STHLHTOY_ID).Visible = False

        Catch ex As SqlException
            MsgBox(ex.ToString)
        Finally
            ' Close connection
            conn.Close()
        End Try
    End Sub
    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' FormEdit("FrmUNIT_MEASURE")
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If lstunit.Items.Count > 0 Then
        '    If MsgBox("Are you sure to delete this Unit Measure ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Sales and Inventory") = MsgBoxResult.Yes Then
        '        sqlSTR = "DELETE FROM TBL_Unit_Measure WHERE ID=" & lstunit.FocusedItem.Text
        '        ExecuteSQLQuery(sqlSTR)
        '        sqlSTR = "SELECT * FROM TBL_Unit_Measure"
        '        FillListView(ExecuteSQLQuery(sqlSTR), lstunit, 0)
        '    End If
        'End If
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class