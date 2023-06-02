Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel


Public Class Form2



    'Create connection
    Dim conn As SqlConnection

    'create data adapter
    Dim da As New SqlDataAdapter

    'create dataset
    Dim ds As DataSet = New DataSet


    'create data adapter
    Dim da2 As SqlDataAdapter

    'create dataset
    Dim ds2 As DataSet = New DataSet

    Dim mRead_Only As Boolean

    Dim GDB As New ADODB.Connection


    'Set up connection string
    Dim cnString As String
    'Public Property PLATOS() As Integer
    '    Get
    '        Return VALUE
    '    End Get
    '    Set(ByVal value As Integer)

    '    End Set
    'End Property

    Private mSTHLH As Integer
    Private mSTHLH2 As Integer

    Private mQUERY As String

    Private f_Alignments(30) As Integer
    Private f_widths(30) As Integer


    Public Property widths(ByVal Index As Integer) As Integer
        Get
            Return f_widths(Index)
        End Get
        Set(ByVal value As Integer)
            f_widths(Index) = value
        End Set
    End Property


    Public Property Alignments(ByVal Index As Integer) As Integer
        Get
            Return f_Alignments(Index)
        End Get
        Set(ByVal value As Integer)
            f_Alignments(Index) = value
        End Set
    End Property



    Public Property STHLHTOY_ID() As Integer
        Get
            Return mSTHLH
        End Get
        Set(ByVal Value As Integer)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            mSTHLH = Value
            'End If
        End Set
    End Property

    Public Property Read_Only() As Boolean
        Get
            Return mRead_Only
        End Get
        Set(ByVal Value As Boolean)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            mRead_Only = Value
            'End If
        End Set
    End Property





    Public Property QUERY_AFTER() As String
        Get
            Return mQUERY
        End Get
        Set(ByVal Value As String)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            mQUERY = Value
            'End If
        End Set
    End Property




    Public Property STHLHONOMATOS_ID() As Integer
        Get
            Return mSTHLH2
        End Get
        Set(ByVal Value As Integer)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            mSTHLH2 = Value
            'End If
        End Set
    End Property



    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GDB.Open(gConnect)


        With Me.GridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        GridView1.AutoSizeRowsMode = _
           DataGridViewAutoSizeRowsMode.AllCells
        paint_ergasies()
        paint_test()


    End Sub

    Public Sub paint_ergasies()
        ' cnString = "Data Source=localhost\SQLEXPRESS;Integrated Security=True;database=MERCURY"
        'Str_Connection = cnString
        cnString = gConSQL


        Dim SQLqry
        SQLqry = Label1.Text '"SELECT NAME,N1,ID FROM ERGATES " ' ORDER BY HME "
        conn = New SqlConnection(cnString)
        Try
            ' Open connection
            conn.Open()

            da = New SqlDataAdapter(SQLqry, conn)

            'create command builder
            Dim cb As SqlCommandBuilder = New SqlCommandBuilder(da)
            ds.Clear()
            'fill dataset
            da.Fill(ds, "PEL")
            GridView1.ClearSelection()
            GridView1.DataSource = ds
            GridView1.DataMember = "PEL"
            'GridView1.Columns(STHLHTOY_ID).Width = 0
            GridView1.Columns(STHLHTOY_ID).Visible = False
            GridView1.Columns(STHLHONOMATOS_ID).Width = 200
            For k = 0 To GridView1.Columns.Count - 1
                If widths(k) > 0 Then
                    GridView1.Columns(k).Width = widths(k)
                End If

                If alignments(k) > 0 Then
                    'GridView1.Columns(k).Width = widths(k)
                    GridView1.Columns(k).DefaultCellStyle.Alignment = alignments(k)
                End If

                'αν θελω σπεσιαλ χρωμα σε ενα κελι
                '   GridView1.Rows(0).Cells(1).Style.BackColor = Color.Red




                'ergates.widths(1) = 100
            Next
            GridView1.RowHeadersWidth = 10 'DEF =43
            GridView1.Columns(2).Frozen = True


            Dim LL As Integer
            LL = GridView1.ColumnCount
            Dim FARDOS As Single
            FARDOS = 0
            Dim J As Integer
            For J = 0 To LL - 2
                FARDOS = FARDOS + GridView1.Columns(J).Width
            Next
            GridView1.Columns(LL - 2).Width = GridView1.Width - FARDOS + 500




            '================================================================================================
            ' πως τυπωνω στην τελευταία σειρά το σύνολο της 4ης στήλης
            'Try
            '    'declaring variable as integer to store the value of the total rows in the datagridview

            '    Dim max As Integer = GridView1.Rows.Count - 1
            '    Dim total As String = "Σύνολο --------->"
            '    Dim tot As Integer = 0
            '    'getting the values of a specific rows
            '    For Each row As DataGridViewRow In GridView1.Rows
            '        'formula for adding the values in the rows
            '        tot += row.Cells(4).Value
            '    Next
            '    GridView1.Rows(max).Cells(4).Value += tot
            '    GridView1.Rows(max).Cells(3).Value = total
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try
            '================================================================================================
















            GridView1.Refresh()







        Catch ex As SqlException
            MsgBox(ex.ToString)
        Finally
            ' Close connection
            conn.Close()
        End Try
    End Sub

    Sub paint_test()



        cnString = gConSQL



        Dim SQLqry
        SQLqry = "select * from YLIKA"  ' Label1.Text '"SELECT NAME,N1,ID FROM ERGATES " ' ORDER BY HME "
        conn = New SqlConnection(cnString)

        conn.Open()

        da = New SqlDataAdapter(SQLqry, conn)

        'create command builder
        Dim cb As SqlCommandBuilder = New SqlCommandBuilder(da)
        ds.Clear()
        'fill dataset
        da.Fill(ds, "PEL")
        GridView1.ClearSelection()
        GridView1.DataSource = ds
        GridView1.DataMember = "PEL"

        Me.Refresh()
    End Sub
End Class