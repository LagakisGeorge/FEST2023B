Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Public Class ergates


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
    Private f_SUMES(30) As Integer


    Public Property widths(ByVal Index As Integer) As Integer
        Get
            Return f_widths(Index)
        End Get
        Set(ByVal value As Integer)
            f_widths(Index) = value
        End Set
    End Property



    Public Property SUMES(ByVal Index As Integer) As Integer
        Get
            Return f_SUMES(Index)
        End Get
        Set(ByVal value As Integer)
            f_SUMES(Index) = value
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




    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' FormAdd("FrmUNIT_MEASURE_ADD")
    End Sub

    'Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    ' End Sub

    Public Sub UPDATEALL()
        Try
            da.Update(ds, "PEL")
            If Len(QUERY_AFTER) > 1 Then
                ExecuteSQLQuery(QUERY_AFTER)
            End If
        Catch ex As Exception
            MsgBox("δεν αποθηκευτηκε" + ex.Message)
        End Try
    End Sub

    Private Sub ergates_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave
        Dim k As Integer
        k = 1
        UPDATEALL()





    End Sub




    'Private Sub AutoSizeRowsMode(ByVal sender As Object, _
    'ByVal e As System.EventArgs) Handles Button7.Click

    '    GridView1.AutoSizeRowsMode = _
    '        DataGridViewAutoSizeRowsMode.AllCells

    'End Sub
    Private Sub FrmUNIT_MEASURE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '  sqlSTR = "SELECT * FROM TBL_Unit_Measure"
        ' FillListView(ExecuteSQLQuery(sqlSTR), lstunit, 0)

        ' If Me.Text = "Προϊόντα     " Then
        'SYNTAGES.Visible = True

        'End If


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
    Public Sub paint_test()



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

                If Alignments(k) > 0 Then
                    'GridView1.Columns(k).Width = widths(k)
                    GridView1.Columns(k).DefaultCellStyle.Alignment = Alignments(k)
                End If

                'αν θελω σπεσιαλ χρωμα σε ενα κελι
                '   GridView1.Rows(0).Cells(1).Style.BackColor = Color.Red




                'ergates.widths(1) = 100
            Next
            GridView1.RowHeadersWidth = 10 'DEF =43
            ' GridView1.Columns(2).Frozen = True


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

            For K = 0 To GridView1.Columns.Count - 1

                If f_SUMES(K) = 1 Then
                    Try
                        'declaring variable as integer to store the value of the total rows in the datagridview

                        Dim max As Integer = GridView1.Rows.Count - 1
                        Dim total As String = "Σύνολο --------->"
                        Dim tot As Integer = 0
                        'getting the values of a specific rows
                        For Each row As DataGridViewRow In GridView1.Rows
                            'formula for adding the values in the rows
                            tot += row.Cells(K).Value
                        Next
                        GridView1.Rows(max).Cells(K).Value += tot
                        GridView1.Rows(max).Cells(0).Value = total
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try

                End If
            Next

            '================================================================================================
















            GridView1.Refresh()







        Catch ex As SqlException
            MsgBox(ex.ToString)
        Finally
            ' Close connection
            conn.Close()
        End Try
    End Sub



    'mia lysh gia to enter
    'Public curcol, currow As Integer

    'Private Sub GridView1_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.CurrentCellChanged
    '    Try
    '        curcol = GridView1.CurrentCell.ColumnIndex
    '        currow = GridView1.CurrentCell.RowIndex
    '    Catch ex As Exception
    '        curcol = 0
    '        currow = 0
    '    End Try
    'End Sub

    'Private Sub GridView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles GridView1.KeyDown
    '    Select Case e.KeyCode
    '        Case Keys.Enter
    '            GridView1.ClearSelection()
    '            Try
    '                If curcol = GridView1.Columns.Count - 2 Then
    '                    If currow < GridView1.Rows.Count - 1 Then
    '                        GridView1.CurrentCell = GridView1(0, currow - 1)
    '                        'original   GridView1.CurrentCell = GridView1(0, currow + 1)
    '                    End If
    '                Else
    '                    GridView1.CurrentCell = GridView1(curcol + 1, currow - 1)
    '                End If
    '            Catch ex As Exception
    '                Exit Try
    '            End Try
    '    End Select
    'End Sub


    '2h lysh gia to enter
    Private Sub GridView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles GridView1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim iCol = GridView1.CurrentCell.ColumnIndex
            Dim iRow = GridView1.CurrentCell.RowIndex
            If iCol = GridView1.Columns.Count - 2 Then  'giati exo krymenh thn kolona toy ID
                If iRow < GridView1.Rows.Count - 1 Then
                    GridView1.CurrentCell = GridView1(0, iRow + 1)
                End If
            Else
                GridView1.CurrentCell = GridView1(iCol + 1, iRow)
            End If
        End If
    End Sub
    'and this will address the "edit" problem mentioned.

    Private Sub GridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellEndEdit
        Dim iCol = GridView1.CurrentCell.ColumnIndex
        Dim iRow = GridView1.CurrentCell.RowIndex
        If iCol = GridView1.Columns.Count - 2 Then 'giati exo krymenh thn kolona toy ID
            If iRow < GridView1.Rows.Count - 1 Then
                GridView1.CurrentCell.Value = GridView1(0, iRow + 1)
            End If
        Else
            If iRow < GridView1.Rows.Count - 1 Then
                SendKeys.Send("{up}")
            End If
            ' GridView1.CurrentCell = GridView1(iCol + 1, iRow)
        End If
    End Sub




    Private Sub GridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Click

        'cnString = "Data Source=localhost\SQLEXPRESS;Integrated Security=True;database=YGEIA"
        ''Str_Connection = cnString
        'Dim C As String


        ''ΠΙΑΝΩ ΤΟ PATIENTID ΤΗΣ ΤΡΕΧΟΥΣΑΣ ΕΓΓΡΑΦΗΣ ΤΟΥ ΑΣΘΕΝΗ
        'C = GridView1.CurrentRow.Cells(0).Value.ToString()




        'Dim SQLqry
        ''SQLqry = "SELECT HME,ERGASIA AS [ΚΑΠΡΟΣ],NEXTOXEIA as [Επ.Οχεία],NEXTGENNA as [Επ.Γέννα],ZVNTA as [Γ.Ζ],NEKRA as [Γ.Ν],MOYMIES AS [ΑΠΩΛ],SEYIOTHESIA AS [ΥΙΟΘ], HMEAPOG AS [ΗΜΕΑΠΟΓ]  , APOGAL AS [AΠOΓ/NA],PARAT as [Παρατηρήσεις],ENOTIO,ID FROM ERGASIES WHERE ENOTIO ='" & Trim(lstCategory.FocusedItem.Text) & "' ORDER BY HME "
        ''"SELECT * FROM ERGASIES "

        'conn = New SqlConnection(cnString)


        ''create data adapter
        ''Dim da2 As SqlDataAdapter

        ''create dataset
        ''Dim ds2 As DataSet = New DataSet






        'Try
        '    ' Open connection
        '    conn.Open()

        '    SQLqry = "SELECT * FROM PATIENTDETAIL WHERE PATIENTID=" + C ' ORDER BY HME "

        '    da2 = New SqlDataAdapter(SQLqry, conn)

        '    'create command builder
        '    Dim cb As SqlCommandBuilder = New SqlCommandBuilder(da2)
        '    ds2.Clear()
        '    'fill dataset
        '    da2.Fill(ds2, "PEL2")
        '    GridView1.ClearSelection()
        '    ' GridView1.Columns.Clear()
        '    'GridView1.Columns(3).Width = 50
        '    GridView1.DataSource = ds2
        '    GridView1.DataMember = "PEL2"
        '    If GridView1.Columns.Count > 0 Then
        '        GridView1.RowHeadersWidth = 40 ' ZONTA
        '    End If

        'Catch ex As SqlException
        '    MsgBox(ex.ToString)
        'Finally
        '    ' Close connection
        '    conn.Close()
        'End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ADD_DETAIL()
    End Sub
    Sub ADD_DETAIL()
        Dim C As String
        'ΠΙΑΝΩ ΤΟ PATIENTID ΤΗΣ ΤΡΕΧΟΥΣΑΣ ΕΓΓΡΑΦΗΣ ΤΟΥ ΑΣΘΕΝΗ
        C = GridView1.CurrentRow.Cells(0).Value.ToString()
        ExecuteSQLQuery("INSERT INTO PATIENTDETAIL (PATIENTID) VALUES (" + C + ")")




    End Sub



    Private Sub GridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub GridView1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            da.Update(ds2, "PEL2")
        Catch ex As Exception
            MsgBox("δεν αποθηκευτηκε" + ex.Message)
        End Try



    End Sub

    'ΑΛΛΑΖΩ ΧΡΩΜΑ ΑΝΑΛΟΓΑ ΜΕ ΤΗΝ ΤΙΜΗ TOY CELL
    '    Use the "RowPostPaint" event
    'The name of the column is NOT the "Header" of the column. You have to go to the properties for the DataGridView => then select the column => then look for the "Name" property
    'I converted this from C# ('From: http://www.dotnetpools.com/Article/ArticleDetiail/?articleId=74)
    'Private Sub GridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
    'Handles dgv_EmployeeTraining.RowPostPaint

    '    If e.RowIndex < Me.dgv_EmployeeTraining.RowCount - 1 Then
    '        Dim dgvRow As DataGridViewRow = Me.dgv_EmployeeTraining.Rows(e.RowIndex)

    '        '<== This is the header Name
    '        'If CInt(dgvRow.Cells("EmployeeStatus_Training_e26").Value) <> 2 Then  


    '        '<== But this is the name assigned to it in the properties of the control
    '        If CInt(dgvRow.Cells("DataGridViewTextBoxColumn15").Value.ToString) <> 2 Then

    '            dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(236, 236, 255)

    '        Else
    '            dgvRow.DefaultCellStyle.BackColor = Color.LightPink

    '        End If

    '    End If

    'End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deleteYLIKA.Click



        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString
        Dim mkod As String = GridView1.CurrentRow.Cells(1).Value.ToString
        Dim ff As New DataTable
        ExecuteSQLQuery("select COUNT(*) FROM SYNTAGES WHERE KODSYNOD='" + mkod + "'", ff)
        If ff.Rows(0)(0) > 0 Then
            MsgBox("AΔYNATH Η ΔΙΑΓΡΑΦΗ , ΓΙΑΤΙ ΣΥΜΜΕΤΕΧΕΙ ΣΕ ΣΥΝΤΑΓΗ")
            Exit Sub
        End If



        ExecuteSQLQuery("DELETE FROM YLIKA WHERE ID=" + mk)

        paint_ergasies()


        'Dim iCol = GridView1.CurrentCell.ColumnIndex
        'Dim iRow = GridView1.CurrentCell.RowIndex
        'If iCol = GridView1.Columns.Count - 2 Then  'giati exo krymenh thn kolona toy ID
        '    If iRow < GridView1.Rows.Count -  1 Then
        '        GridView1.CurrentCell = GridView1(0, iRow + 1)
        '    End If
        'Else
        '    GridView1.CurrentCell = GridView1(iCol + 1, iRow)
        'End If



    End Sub


    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim filename As String = "c:\mercvb\ektyp.xlsx"
        Dim sheetname As String = "Φύλλο1"
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xl As Excel.Worksheet
        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add 'αν ηθελα να το ανοιξω αντι για add -> Open(filename)
        xlWorkBook.Worksheets.Add()  '(1)
        xl = xlWorkBook.Worksheets(1) ' .Add
        xlApp.Visible = True  'ΜΠΟΡΩ ΝΑ ΤΟ ΒΛΕΠΩ
        xl.Name = "HL07"
        Dim WS(30) As Microsoft.Office.Interop.Excel.Worksheet

        Dim dt As New DataTable
        Dim k As Integer
        Dim mn1 As String = "1"
        Dim sql As String '= "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2,ID  FROM YLIKA WHERE N1=" + mn1 + " ORDER BY KOD "



        sql = Label1.Text


        ExecuteSQLQuery(sql, dt) 'D.PATIENTID,CHMEEIS desc

        xl.Cells(1, 2).value = Me.Text  '"Αρχείο Υλικών"  '"ΕΠΙΚΕΦΑΛΙΔΑ" + "EIS"

        Dim sken As Single = 0
        Dim seopy As Single = 0

        Dim mSeir As Integer = 2

        Dim L As Integer

        For L = 0 To dt.Columns.Count - 1
            xl.Cells(mSeir, L + 1).value = dt.Columns(L).Caption 'a
        Next



0:
        mSeir = 2
        For k = 0 To dt.Rows.Count - 1
            mSeir = mSeir + 1
            For L = 0 To dt.Columns.Count - 1
                xl.Cells(mSeir, L + 1).value = dt.Rows(k)(L)  'aa
            Next
        Next

        xl.Columns.AutoFit()
        xlApp.Visible = True
        mreleaseObject(xlApp)
        mreleaseObject(xlWorkBook)
        mreleaseObject(xl)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SYNTAGES.Click
        'συνταγες
        Dim mergates As New ergates()
        For k = 0 To 20
            mergates.widths(7) = 100
        Next
        Dim Mn1 As String
        'Mn1 = GridView1(GridView1.CurrentRow)(GridView1.CurrentCell).tostring    '  Split(KATHG.Text, ";")(0)
        mergates.Text = "Συνταγή " + GridView1.CurrentRow.Cells(0).Value.ToString
        'mergates.Label1.Text = "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2,ID  FROM YLIKA WHERE N1=" + Mn1 + " ORDER BY KOD "

        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(1).Value.ToString


        If Trim(mk.Length) = 0 Then
            Exit Sub
        End If

        mergates.Label1.Text = "SELECT KOD ,KODSYNOD ,POSOSTO,(SELECT top 1 ONO FROM YLIKA WHERE KOD=SYNTAGES.KODSYNOD) AS [ΠΕΡΙΓΡΑΦΗ] ,ID  FROM SYNTAGES  where KOD='" + mk + "'"

        'mergates.Label1.Text = "SELECT KOD AS [ΚΩΔ],KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],ID  FROM SYNTAGES  where KOD='" + mk + "'"

        ' ergates.MdiParent = Me
        mergates.WindowState = FormWindowState.Maximized
        mergates.STHLHONOMATOS_ID = 0
        mergates.STHLHTOY_ID = 4
        mergates.widths(1) = 100
        mergates.QUERY_AFTER = ""  '"update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            mergates.widths(KK) = 100
        Next
        mergates.Label2.Text = "synt"  'per '"υλικα...."   ' KATHG.Text
        mergates.widths(0) = 400
        gMenu = 22
        mergates.Read_Only = False

        mergates.ShowDialog()




        'mergates.TopLevel = False
        'mergates.Visible = True
        'mergates.FormBorderStyle = FormBorderStyle.None
        'mergates.Dock = DockStyle.Fill
        'Dim PAGE As New TabPage
        'Dim N As Integer = TabControl11.TabPages.Count
        'PAGE.Text = per  '"ΥΛΙΚΑ.....   ."
        'TabControl11.TabPages.Add(PAGE)

        'mergates.Width = TabControl11.Width
        'mergates.Height = TabControl11.Height
        'TabControl11.TabPages(N).Controls.Add(mergates)
        'TabControl11.SelectTab(N)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ANALYTICS.Click

        Dim mergates As New ergates()
        For k = 0 To 20
            mergates.widths(7) = 100
        Next
        Dim Mn1 As String
        'Mn1 = GridView1(GridView1.CurrentRow)(GridView1.CurrentCell).tostring    '  Split(KATHG.Text, ";")(0)
        mergates.Text = "Τιμολόγια"
        'mergates.Label1.Text = "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2,ID  FROM YLIKA WHERE N1=" + Mn1 + " ORDER BY KOD "

        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString


        If Trim(mk.Length) = 0 Then
            Exit Sub
        End If

        mergates.Label1.Text = "SELECT T.N2 ,T.HME ,T.POSO,T.ID  FROM TIMSANAL T LEFT JOIN PARTIDES P ON T.IDPART=P.ID where IDTIMS='" + mk + "'"

        'mergates.Label1.Text = "SELECT KOD AS [ΚΩΔ],KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],ID  FROM SYNTAGES  where KOD='" + mk + "'"

        ' ergates.MdiParent = Me
        mergates.WindowState = FormWindowState.Maximized
        mergates.STHLHONOMATOS_ID = 0
        mergates.STHLHTOY_ID = 3
        mergates.widths(1) = 100
        mergates.QUERY_AFTER = ""  '"update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            mergates.widths(KK) = 100
        Next
        mergates.Label2.Text = "synt"  'per '"υλικα...."   ' KATHG.Text
        mergates.widths(0) = 400
        gMenu = 22

        mergates.Read_Only = True
        mergates.ShowDialog()


    End Sub


    Private Sub AnalPartidas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AnalPartidas.Click

        Dim mergates As New ergates()
        For k = 0 To 20
            mergates.widths(7) = 100
        Next
        Dim Mn1 As String
        'Mn1 = GridView1(GridView1.CurrentRow)(GridView1.CurrentCell).tostring    '  Split(KATHG.Text, ";")(0)
        mergates.Text = "Πωλήσεις Παρτίδας"
        'mergates.Label1.Text = "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2,ID  FROM YLIKA WHERE N1=" + Mn1 + " ORDER BY KOD "

        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(0).Value.ToString


        If Trim(mk.Length) = 0 Then
            Exit Sub
        End If

        mergates.Label1.Text = "SELECT (SELECT TOP 1 EPO FROM PEL WHERE EIDOS='e' AND KOD=TIMSPOL.PROM) AS [ΠΕΛΑΤΗΣ],ATIM ,HME ,POSO AS [ΠΟΣΟΤΗΤΑ],PARTIDA AS [ΠΑΡΤΙΔΑ],ID  FROM TIMSPOL  where PARTIDA='" + mk + "'"

        'mergates.Label1.Text = "SELECT KOD AS [ΚΩΔ],KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],ID  FROM SYNTAGES  where KOD='" + mk + "'"

        ' ergates.MdiParent = Me
        mergates.WindowState = FormWindowState.Maximized
        mergates.STHLHONOMATOS_ID = 0
        mergates.STHLHTOY_ID = 5
        mergates.widths(1) = 100
        mergates.QUERY_AFTER = ""  '"update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            mergates.widths(KK) = 100
        Next
        mergates.Label2.Text = "synt"  'per '"υλικα...."   ' KATHG.Text
        mergates.widths(0) = 200
        gMenu = 22
        mergates.Read_Only = True

        mergates.ShowDialog()

    End Sub

    'Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Proeleysi_Partidas.Click
    '    Dim mergates As New ergates()
    '    For k = 0 To 20
    '        mergates.widths(7) = 100
    '    Next
    '    Dim Mn1 As String
    '    'Mn1 = GridView1(GridView1.CurrentRow)(GridView1.CurrentCell).tostring    '  Split(KATHG.Text, ";")(0)
    '    mergates.Text = "Προέλευση Παρτίδας"
    '    'mergates.Label1.Text = "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2,ID  FROM YLIKA WHERE N1=" + Mn1 + " ORDER BY KOD "

    '    ' ο κωδικος του προιοντος που διαλεξα
    '    Dim mk As String = GridView1.CurrentRow.Cells(0).Value.ToString


    '    If Trim(mk.Length) = 0 Then
    '        Exit Sub
    '    End If

    '    mergates.Label1.Text = "SELECT PROM,ATIM ,HME ,POSO,PARTIDA,ID  FROM TIMSANAL  where PARTIDA='" + mk + "'"

    '    'mergates.Label1.Text = "SELECT KOD AS [ΚΩΔ],KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],ID  FROM SYNTAGES  where KOD='" + mk + "'"

    '    ' ergates.MdiParent = Me
    '    mergates.WindowState = FormWindowState.Maximized
    '    mergates.STHLHONOMATOS_ID = 0
    '    mergates.STHLHTOY_ID = 5
    '    mergates.widths(1) = 100
    '    mergates.QUERY_AFTER = ""  '"update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
    '    For KK As Integer = 0 To 6
    '        mergates.widths(KK) = 100
    '    Next
    '    mergates.Label2.Text = "synt"  'per '"υλικα...."   ' KATHG.Text
    '    mergates.widths(0) = 200
    '    gMenu = 22
    '    mergates.Read_Only = True

    '    mergates.ShowDialog()

    'End Sub

    Private Sub Proeleysi_Partidas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Proeleysi_Partidas.Click
        Dim mergates As New ergates()
        For k = 0 To 20
            mergates.widths(7) = 100
        Next
        Dim Mn1 As String
        Dim mkod = GridView1.CurrentRow.Cells(2).Value.ToString
        Dim mono = GridView1.CurrentRow.Cells(3).Value.ToString
        Dim mYpol As String = GridView1.CurrentRow.Cells(6).Value.ToString
        'Mn1 = GridView1(GridView1.CurrentRow)(GridView1.CurrentCell).tostring    '  Split(KATHG.Text, ";")(0)
        mergates.Text = mkod + " " + mono + " Προέλευση Παρτίδας"
        'mergates.Label1.Text = "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2,ID  FROM YLIKA WHERE N1=" + Mn1 + " ORDER BY KOD "

        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString


        If Trim(mk.Length) = 0 Then
            Exit Sub
        End If

        mergates.Label1.Text = "SELECT A.CH1 AS [ΚΩΔ.ΥΛΙΚΟΥ],A.HME AS [HM.ΠΑΡΑΓ],CH2 AS [ΩΡΑ ΠΑΡΑΓ] ,A.POSO AS [ΠΟΣΟΤΗΤΑ],T.ATIM AS [ΑΡ.ΤΙΜΟΛ],T.HME AS [ΗΜ.ΤΙΜΟΛ],T.PROM AS [ΠΡΟΜΗΘ],A.ID ,A.IDTIMS,A.N2,IDPART  FROM TIMSANAL A LEFT JOIN TIMS T ON A.IDTIMS=T.ID  where A.N2=" + GridView1.CurrentRow.Cells(0).Value.ToString    '"IDPART=" + mk + ""
        Dim fofo As New DataTable
        Dim isApografh As Integer = 0
        ExecuteSQLQuery("SELECT A.CH1 AS [ΚΩΔ.ΥΛΙΚΟΥ],A.HME AS [HM.ΠΑΡΑΓ],CH2 AS [ΩΡΑ ΠΑΡΑΓ] ,A.POSO AS [ΠΟΣΟΤΗΤΑ],T.ATIM AS [ΑΡ.ΤΙΜΟΛ],T.HME AS [ΗΜ.ΤΙΜΟΛ],T.PROM AS [ΠΡΟΜΗΘ],A.ID ,A.IDTIMS,A.N2,IDPART  FROM TIMSANAL A LEFT JOIN TIMS T ON A.IDTIMS=T.ID  where A.N2=" + GridView1.CurrentRow.Cells(0).Value.ToString, fofo)


        ' ExecuteSQLQuery("SELECT DISTINCT HME AS [HM.ΠΑΡΑΓ],CH2 AS [ΩΡΑ ΠΑΡΑΓ] ,TEMAX AS [ΠΟΣΟΤΗΤΑ] FROM TIMSANAL  where N2=" + GridView1.CurrentRow.Cells(0).Value.ToString, fofo)


        If fofo.Rows.Count = 0 Then
            mergates.Label1.Text = "select PARTIDA AS [ΠΑΡΤΙΔΑ],HME as [ΗΜΕΡ.ΠΑΡΑΓ],KOD as [ΠΡΟΪΟΝ],TEMAXIA AS [ΠΑΡΑΧΘΕΝΤΑ],TEMAXIA-YPOL AS [ΠΩΛΗΘΕΝΤΑ],YPOL AS [ΥΠΟΛΟΙΠΟ],'' AS [ ],ID  FROM PARTIDES where ID= " + mk
            isApografh = 1
        End If


        '"IDPART=" + mk + "")









        'mergates.Label1.Text = "SELECT KOD AS [ΚΩΔ],KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],ID  FROM SYNTAGES  where KOD='" + mk + "'"

        ' ergates.MdiParent = Me
        mergates.WindowState = FormWindowState.Maximized
        mergates.STHLHONOMATOS_ID = 0
        mergates.STHLHTOY_ID = 7
        mergates.widths(1) = 100
        mergates.QUERY_AFTER = ""  '"update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            mergates.widths(KK) = 100
        Next
        mergates.Label2.Text = mkod + " " + mono + " Προέλευση Παρτίδας ;" + MYPOL  '  "synt"  'per '"υλικα...."   ' KATHG.Text
        mergates.widths(0) = 200
        gMenu = 22
        mergates.Read_Only = True
        If isApografh = 1 Then
            mergates.STHLHTOY_ID = 7
            mergates.DELETE_APOG_PART.Visible = True

        Else
            mergates.delete_label.Visible = True
            mergates.CMDEPANEKTYPOSI.Visible = True
        End If


        mergates.ShowDialog()
    End Sub


    Private Function FindInGrid(ByRef dgvGrid As DataGridView, ByVal strFind As String, Optional ByVal bStartAtBeggining As Boolean = False) As Point

        'call it like:
        '    ' forces a start at the grid cell 0,0
        'FindInGrid(Me.dgvMyGridView, Me.m_sFind, True)

        '    ' starts one cell beyond the staic member X, if X = last column then increments static member Y and sets X = 0
        'FindInGrid(Me.dgvMyGridView, Me.m_sFind)

        Dim bFound As Boolean = False
        Static pResult As Point
        Static X As Integer
        Static Y As Integer
        Dim dgvCell As DataGridViewCell
        Dim dgvRow As DataGridViewRow = Nothing
        For Each dgvRow In dgvGrid.SelectedRows()
            dgvRow.Selected = False
        Next
        If bStartAtBeggining Then
            pResult.X = 0
            pResult.Y = 0
        Else
            If pResult.X < dgvGrid.Columns.Count - 1 Then
                pResult.X += 1
            Else
                pResult.X = 0
                If pResult.Y < dgvGrid.Rows.Count - 1 Then
                    pResult.Y += 1
                Else
                    pResult.Y = 0
                End If
            End If
        End If
        For Y = pResult.Y To dgvGrid.Rows.Count - 1
            For X = pResult.X To dgvGrid.Columns.Count - 1
                dgvCell = dgvGrid(X, Y)
                If Not IsDBNull(dgvCell.Value) Then
                    If Not dgvCell.Value = Nothing Then
                        If dgvCell.Value.ToString.ToLower.Contains(strFind.ToLower) Then
                            If Not dgvCell.Value = Nothing Then
                                pResult.X = X
                                pResult.Y = Y
                                bFound = True
                                Exit For
                            End If
                        End If
                    End If
                End If


            Next
            If bFound Then
                Exit For
            Else
                pResult.X = 0
            End If
            If Y = dgvGrid.Rows.Count - 1 Then
                pResult.Y = 0
                If dgvGrid.CurrentCell.Value.ToString.ToLower.Contains(strFind.ToLower) Then
                    MessageBox.Show("δεν υπάρχει αλλο " & strFind & " για εύρεση.", _
                    "Grid Search...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Else
                    MessageBox.Show("'" & strFind & "' δεν βρέθηκε ", _
                    "Grid Search...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                End If
            End If
        Next
        If bFound Then
            Try
                dgvGrid.Item(pResult.X, pResult.Y).Selected = True
                dgvGrid.CurrentCell = dgvGrid.Item(pResult.X, pResult.Y)

            Catch ex As Exception

            End Try

        End If
    End Function

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim c As String = InputBox("Αναζήτηση ", , )
        FindInGrid(Me.GridView1, cFind.Text)
    End Sub

    Private Sub delete_label_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles delete_label.Click

        Dim mpART As String = GridView1.CurrentRow.Cells("ID").Value.ToString
        ' Dim mORA As String = GridView1.CurrentRow.Cells("ΩΡΑ ΠΑΡΑΓ").Value.ToString
        ExecuteSQLQuery("delete from PEL WHERE ID=" + mpART + " AND (HOTELID=0 OR HOTELID IS NULL)")
        paint_ergasies()

    End Sub

    

    Private Sub DEL_TIMAGOR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DEL_TIMAGOR.Click

        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString
        Dim mkod As String = GridView1.CurrentRow.Cells(1).Value.ToString
        Dim ff As New DataTable
        ExecuteSQLQuery("select POSO-YPOL AS DIF FROM TIMS WHERE ID=" + mk, ff)



        'If ff.Rows(0)(0) > 0 Then
        'MsgBox("AΔYNATH Η ΔΙΑΓΡΑΦΗ , ΓΙΑΤΙ EXEI ΣΥΜΜΕΤOXH ΣΕ ΠΑΡΤΙΔΑ")
        'Exit Sub
        'End If



        ExecuteSQLQuery("DELETE FROM TIMS WHERE ID=" + mk)

        paint_ergasies()

    End Sub

    Private Sub GridView1_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellContentClick

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Dim k As Integer
        k = 1

        If Read_Only = False Then
            UPDATEALL()
        End If

        'MDIMain.cmdDelete.Enabled = False
        Me.Close()
    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        Label1.Text = Replace(Label1.Text, "top 400", " ")

        paint_ergasies()
    End Sub

    Private Sub cmdAdd_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        FrmAddEidos.n1.Text = n1.Text
        FrmAddEidos.IsNew = True
        FrmAddEidos.ShowDialog()
        paint_ergasies()

    End Sub

    Private Sub deleteTIMPOL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deleteTIMPOL.Click

        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString
        Dim mkod As String = GridView1.CurrentRow.Cells(1).Value.ToString
        Dim ff As New DataTable
        'ExecuteSQLQuery("select * AS DIF FROM TIMSPOL WHERE ID=" + mk, ff)


        ExecuteSQLQuery("select *  FROM TIMSPOL WHERE ID=" + mk, ff)
        Dim MP As String = ff.Rows(0)("PARTIDA").ToString
        ' MsgBox("AΔYNATH Η ΔΙΑΓΡΑΦΗ , ΓΙΑΤΙ EXEI ΣΥΜΜΕΤOXH ΣΕ ΠΑΡΤΙΔΑ")
        ' Exit Sub
        'End If

        ' ========================  PROSOXH ========================================
        'ΣΤΑ ΠΙΣΤΩΤΙΚΑ ΣΤΟ TIMSPOL ΑΠΟΘΗΚΕΥΩ ΤΑ ΕΞΗΣ ΓΙΑ ΝΑ ΜΠΟΡΩ ΝΑ ΤΑ ΣΒΗΣΩ
        'C1='ΠΙΣ'
        'Ν1= ID PARTIDES
        'N2 = ID TIMS 





        ' Exit Sub






        Dim ans2 As Integer
        ans2 = MsgBox("Να διαγραφεί το τιμολόγιο " + ff.Rows(0)(2).ToString, MsgBoxStyle.YesNo)

        If ans2 = vbNo Then
            Exit Sub
        End If




        If ff.Rows(0)("C1").ToString = "ΠΙΣ" Then

            If Val(MP) > 0 Then  'ΗΤΑΝ ΕΤΟΙΜΟ ΕΙΧΕ ΠΑΡΤΙΔΑ

                ExecuteSQLQuery("UPDATE PARTIDES SET YPOL=YPOL-" + Replace(ff.Rows(0)("POSO").ToString, ",", ".") + "  WHERE ID=" + ff(0)("N1").ToString)
                ExecuteSQLQuery("DELETE FROM TIMSPOL WHERE ID=" + mk)

            Else   'ΗΤΑΝ ΥΛΙΚΟ ΔΕΝ ΕΙΧΕ ΠΑΡΤΙΔΑ ΠΑΩ TIMSANAL->TIMS ΓΙΑ ΝΑ ΣΒΗΣΩ ΤΗΝ ΠΟΣΟΤΗΤΑ

                ExecuteSQLQuery("DELETE FROM TIMS WHERE ID=" + ff(0)("N1").ToString)
                ExecuteSQLQuery("DELETE FROM TIMSPOL WHERE ID=" + mk)



            End If

            paint_ergasies()
            MsgBox("διεγράφη")

            Exit Sub
        End If







        If Val(MP) > 0 Then  'ΗΤΑΝ ΕΤΟΙΜΟ ΕΙΧΕ ΠΑΡΤΙΔΑ

            ExecuteSQLQuery("UPDATE PARTIDES SET YPOL=YPOL+" + Replace(ff.Rows(0)("POSO").ToString, ",", ".") + "  WHERE PARTIDA=" + MP)


            ExecuteSQLQuery("DELETE FROM TIMSPOL WHERE ID=" + mk)

        Else   'ΗΤΑΝ ΥΛΙΚΟ ΔΕΝ ΕΙΧΕ ΠΑΡΤΙΔΑ ΠΑΩ TIMSANAL->TIMS ΓΙΑ ΝΑ ΣΒΗΣΩ ΤΗΝ ΠΟΣΟΤΗΤΑ

            Dim mH As String = ff.Rows(0)("HME").ToString

            Dim FOFI As New DataTable
            ExecuteSQLQuery("select *  FROM TIMSANAL WHERE N2=" + mk, FOFI)

            For L As Integer = 0 To FOFI.Rows.Count - 1
                Dim MID As String = FOFI.Rows(L)("ID").ToString




                GDB.Execute("UPDATE TIMS SET YPOL=YPOL+" + Replace(FOFI.Rows(L)("POSO").ToString, ",", ".") + " WHERE ID=" + FOFI.Rows(L)("IDTIMS").ToString)


                GDB.Execute("DELETE FROM TIMSANAL WHERE ID=" + MID)



            Next

            ExecuteSQLQuery("DELETE FROM TIMSPOL WHERE ID=" + mk)

        End If

        paint_ergasies()
        MsgBox("διεγράφη")
    End Sub

    Private Sub DELETE_APOG_PART_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETE_APOG_PART.Click
        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString
        Dim FOFO As New DataTable
        ExecuteSQLQuery("DELETE FROM PARTIDES WHERE ID=" + mk, FOFO)
        MsgBox("ΔΙΕΓΡΑΦΗ")
        paint_ergasies()

    End Sub

    Private Sub kinhseis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kinhseis.Click
        Dim mergates As New ergates()
        For k = 0 To 20
            mergates.widths(7) = 100
        Next
        Dim Mn1 As String
        'Mn1 = GridView1(GridView1.CurrentRow)(GridView1.CurrentCell).tostring    '  Split(KATHG.Text, ";")(0)
        mergates.Text = "Πωλήσεις Παρτίδας"
        'mergates.Label1.Text = "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2,ID  FROM YLIKA WHERE N1=" + Mn1 + " ORDER BY KOD "

        ' ο κωδικος του προιοντος που διαλεξα
        Dim mk As String = GridView1.CurrentRow.Cells(0).Value.ToString


        If Trim(mk.Length) = 0 Then
            Exit Sub
        End If




        Dim cc As String

        Dim TYPOS As String = ""
        Dim DTT As New DataTable
        ExecuteSQLQuery("SELECT * FROM YLIKA WHERE KOD='" + mk + "'", DTT)
        If DTT.Rows.Count = 0 Then
            MsgBox("ΔΕΝ ΒΡΙΣΚΩ ΤΟ ΥΛΙΚΟ")
            Exit Sub
        Else
            If DTT.Rows(0)("N1") = 4 Then
                TYPOS = "PRO"
                'ΚΙΝΗΣΕΙΣ ΤΟΥ ΕΙΔΟΥΣ
                ' cc = "SELECT  'TIMSANAL' AS TYPE, POSO,HME,N2 AS PARTIDA,ID , '' AS CC FROM TIMSANAL  WHERE LEFT(CH1,6)='" + mk + "'  "
                cc = " SELECT 'ΑΓΟΡ-ΑΠΟΓ' AS TYPE,  YPOL AS [ΥΠΟΛ],HME,'' AS [ΠΑΡΤΙΔΑ]  ,ATIM AS [ΑΡ.ΤΙΜ],POSO AS [ΑΡΧ.ΠΟΣΟΤ],POSO-YPOL AS [ΑΝΑΛ],ID FROM TIMS WHERE KOD='" + mk + "'  "
                cc = cc + " UNION  SELECT 'ΠΑΡΤΙΔΕΣ' AS TYPE, YPOL AS [ΥΠΟΛ],HME,PARTIDA AS [PARTIDA] ,'' AS [ΑΡ.ΤΙΜ],TEMAXIA AS [ΑΡΧ.ΠΟΣΟΤ],TEMAXIA-YPOL AS [ΑΝΑΛ],ID FROM PARTIDES WHERE KOD='" + mk + "'  "

                cc = cc + " UNION SELECT  'ΤΙΜ.ΠΩΛ' AS TYPE, 0 AS [ΥΠΟΛ],HME, PARTIDA AS [ΠΑΡΤΙΔΑ],ATIM AS [ΑΡ.ΤΙΜ] ,POSO AS [ΑΡΧ.ΠΟΣΟΤ],POSO AS [ΑΝΑΛ],ID FROM TIMSPOL  WHERE KOD='" + mk + "'"
                cc = cc + "UNION SELECT  'ΚΙΝΗΣ.ΑΝΑΛ' AS TYPE, 0  AS [ΥΠΟΛ],HME,N2 AS [ΠΑΡΤΙΔΑ],'' AS [ΑΡ.ΤΙΜ],POSO AS [ΑΡΧ.ΠΟΣΟΤ],POSO AS [ΑΝΑΛ],ID  FROM TIMSANAL  WHERE LEFT(CH1,6)='" + mk + "'  "

            ElseIf DTT.Rows(0)("N1") = 1 Then
                cc = " SELECT 'ΑΓΟΡ-ΑΠΟΓ' AS TYPE,  YPOL AS [ΥΠΟΛ],HME,'' AS [ΠΑΡΤΙΔΑ]  ,ATIM AS [ΑΡ.ΤΙΜ],POSO AS [ΑΡΧ.ΠΟΣΟΤ] ,POSO-YPOL AS [ΑΝΑΛ],ID FROM TIMS WHERE KOD='" + mk + "'  "
                cc = cc + " UNION SELECT  'TIM.ΠΩΛ' AS TYPE, 0 AS [ΥΠΟΛ],HME, PARTIDA  ,ATIM AS [ΑΡ.ΤΙΜ] ,POSO AS [ΑΡΧ.ΠΟΣΟΤ],POSO AS [ΑΝΑΛ],ID FROM TIMSPOL  WHERE KOD='" + mk + "'"



                TYPOS = "AYL"
            ElseIf DTT.Rows(0)("N1") = 2 Then
                TYPOS = "EMP"
                cc = " SELECT 'ΑΓΟΡ-ΑΠΟΓ' AS TYPE,  YPOL AS [ΥΠΟΛ],HME,'' AS [ΠΑΡΤΙΔΑ] ,ATIM AS [ΑΡ.ΤΙΜ],POSO AS [ΑΡΧ.ΠΟΣΟΤ],POSO-YPOL AS [ΑΝΑΛ], ID FROM TIMS WHERE KOD='" + mk + "'  "
                cc = cc + " UNION SELECT  'TIM.ΠΩΛ' AS TYPE, 0 AS [ΥΠΟΛ],HME, PARTIDA ,ATIM AS [ΑΡ.ΤΙΜ] ,POSO AS [ΑΡΧ.ΠΟΣΟΤ],POSO AS [ΑΝΑΛ],ID FROM TIMSPOL  WHERE KOD='" + mk + "'"
            Else

                cc = " SELECT 'ΑΓΟΡ-ΑΠΟΓ' AS TYPE,  YPOL AS [ΥΠΟΛ],HME,'' AS [ΠΑΡΤΙΔΑ]  ,ATIM AS [ΑΡ.ΤΙΜ],POSO AS [ΑΡΧ.ΠΟΣΟΤ],POSO-YPOL AS [ΑΝΑΛ],ID FROM TIMS WHERE KOD='" + mk + "'  "
                cc = cc + " UNION SELECT  'TIM.ΠΩΛ' AS TYPE, 0 AS [ΥΠΟΛ],HME, PARTIDA  ,ATIM AS [ΑΡ.ΤΙΜ] ,POSO AS [ΑΡΧ.ΠΟΣΟΤ],POSO-YPOL AS [ΑΝΑΛ],ID FROM TIMSPOL  WHERE KOD='" + mk + "'"





                TYPOS = "ANA"
            End If
        End If









        ''ΚΙΝΗΣΕΙΣ ΤΟΥ ΕΙΔΟΥΣ
        'cc = "SELECT  'TIMSANAL' AS TYPE, POSO,HME,N2 AS PARTIDA,ID , '' AS CC FROM TIMSANAL  WHERE LEFT(CH1,6)='" + mk + "'  "
        'cc = cc + " UNION  SELECT 'TIMS' AS TYPE,  POSO,HME,'' AS PARTIDA,ID ,ATIM AS CC FROM TIMS WHERE KOD='" + mk + "'  "
        'cc = cc + " UNION  SELECT 'PARTIDES' AS TYPE, YPOL AS POSO,HME,PARTIDA,ID ,'' AS CC FROM PARTIDES WHERE KOD='" + mk + "'  "

        'cc = cc + " UNION SELECT  'TIMSPOL' AS TYPE, POSO,HME, PARTIDA ,ID ,ATIM AS CC  FROM TIMSPOL  WHERE KOD='" + mk + "'"

        mergates.Alignments(1) = DataGridViewContentAlignment.MiddleRight
        mergates.Alignments(5) = DataGridViewContentAlignment.MiddleRight
        mergates.Alignments(6) = DataGridViewContentAlignment.MiddleLeft

        mergates.Label1.Text = cc  ' "SELECT (SELECT TOP 1 EPO FROM PEL WHERE EIDOS='e' AND KOD=TIMSPOL.PROM) AS [ΠΕΛΑΤΗΣ],ATIM ,HME ,POSO AS [ΠΟΣΟΤΗΤΑ],PARTIDA AS [ΠΑΡΤΙΔΑ],ID  FROM TIMSPOL  where PARTIDA='" + mk + "'"

        'mergates.Label1.Text = "SELECT KOD AS [ΚΩΔ],KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],ID  FROM SYNTAGES  where KOD='" + mk + "'"

        ' ergates.MdiParent = Me
        mergates.WindowState = FormWindowState.Maximized
        mergates.STHLHONOMATOS_ID = 0
        mergates.STHLHTOY_ID = 7
        mergates.widths(1) = 100
        mergates.QUERY_AFTER = ""  '"update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            mergates.SUMES(KK) = 0
            mergates.widths(KK) = 100
        Next
        mergates.SUMES(1) = 1
        mergates.Label2.Text = DTT.Rows(0)("KOD").ToString + "  " + DTT.Rows(0)("ONO").ToString 'per '"υλικα...."   ' KATHG.Text
        mergates.widths(5) = 80
        gMenu = 22
        mergates.Read_Only = True

        mergates.ShowDialog()



    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMDEPANEKTYPOSI.Click

        ' ο κωδικος του προιοντος που διαλεξα
        ' Dim mID As String = GridView1.CurrentRow.Cells("IDTIMS").Value.ToString
        'Dim mpos As String '= GridView1.CurrentRow.Cells("IDTIMS").Value.ToString
        Dim mpART As String = GridView1.CurrentRow.Cells("N2").Value.ToString
        Dim mHMERA As String = GridView1.CurrentRow.Cells(1).Value.ToString
        Dim mORA As String = GridView1.CurrentRow.Cells("ΩΡΑ ΠΑΡΑΓ").Value.ToString

        Dim L1 As Integer = InStr(mHMERA, " ")
        If L1 = 0 Then L1 = 10
        mHMERA = Mid(mHMERA, 1, L1)


        'mpos = GridView1.CurrentRow.Cells("ΠΟΣΟΤΗΤΑ").Value.ToString
        'SELECT A.CH1 AS [ΚΩΔ.ΥΛΙΚΟΥ],A.HME AS [HM.ΠΑΡΑΓ],CH2 AS [ΩΡΑ ΠΑΡΑΓ] ,A.POSO AS [ΠΟΣΟΤΗΤΑ],T.ATIM AS [ΑΡ.ΤΙΜΟΛ],T.HME AS [ΗΜ.ΤΙΜΟΛ],T.PROM AS [ΠΡΟΜΗΘ],A.ID,A.IDTIMS,N2 ,IDPART FROM TIMSANAL A INNER JOIN TIMS T ON A.IDTIMS=T.ID  where A.N2=" + GridView1.CurrentRow.Cells(0).Value.ToString    '"IDPART=" + mk + ""

        'παιρνω ολες τις σειρες που εχουν την ιδια ωρα
        'αυξανω στο tims ΤΟ ΥΠΟΛΟΙΠΟ   ( ΜΕ ΤΗΝ ΒΟΗΘΕΙΑ ΤΟΥ Α.IDTIMS )





        Dim mProion As String = Mid(Label2.Text, 1, 6)

        Dim R As New DataTable


        ExecuteSQLQuery("SELECT POSO,T.CH1 AS KODIKOS,T.N2 AS PART,Y.ONO,T.TEMAX FROM TIMSANAL T INNER JOIN YLIKA Y ON T.CH1=Y.KOD WHERE T.N2=" + mpART + " AND T.CH2='" + mORA + "'", R)

        If IsDBNull(R(0)("TEMAX")) Then

            MsgBox(" αδυνατη η επανεκτύπωση. Δεν υπάρχουν τα τεμάχια ")
            Exit Sub


        End If

        Dim TELBARCODE As String
        TELBARCODE = "'0205200000" + Mid(mProion + Space(6), 1, 6) + "300" + Mid(Str(100000 + R(0)("TEMAX")), 3, 5) + "10000" + Mid(LTrim(R(0)("PART").ToString), 1, 6)




        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\test-label-epan.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sh1")
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value
        xlWorkSheet.Cells(7, 2) = Mid(Label2.Text, 8, 40) ' R(0)("ONO").ToString  '"" 'onomaProion
        xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + R(0)("TEMAX").ToString  'minTem)
        xlWorkSheet.Cells(9, 1) = "ΠΑΡΤΙΔΑ/LOT :" + Str(R(0)("PART").ToString) ' MP)   N2
        xlWorkSheet.Cells(15, 1) = TELBARCODE
        xlWorkSheet.Cells(18, 2) = mProion 'R(0)("KODIKOS").ToString ' kodPROION
        xlWorkSheet.Cells(8, 4) = mHMERA + " " + mORA 'R(0)("KODIKOS").ToString ' kodPROION

        'Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)
        xlWorkBook.Save()

        xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

















        'GDB.BeginTrans()
        'Try


        '    For N As Integer = 0 To R.Rows.Count - 1
        '        Dim poso As String = R.Rows(N)("poso").ToString
        '        Dim idT As String = R.Rows(N)("IDTIMS").ToString
        '        Dim id As String = R.Rows(N)("ID").ToString
        '        Dim idPART As String = R.Rows(N)("IDPART").ToString
        '        ' ΣΕ ΠΡΟΦΟΡΜΕΣ
        '        'IDTIMS -> ΒΑΖΩ ΤΟ ID THΣ ΠΑΤΡΙΚΗΣ PARTIDAS    N2=ΑΡΙΘΜΟΣ ΠΑΡΤΙΔΑΣ ΠΑΡΑΧΘΕΙΣΑΣ IDPART=-1
        '        ' ΣΕ ΚΑΝΟΝΙΚΕΣ ΠΑΡΤΙΔΟΠΟΙΗΣΕΙΣ ( ΤΙΜΟΛΟΓΙΑ )
        '        'IDTIMS -> ΒΑΖΩ ΤΟ ID TOY ΠΑΤΡΙΚOY TIMOLOGIOY  N2=ΑΡΙΘΜΟΣ ΠΑΡΤΙΔΑΣ ΠΑΡΑΧΘΕΙΣΑΣ IDPART=0

        '        'ENHMERVNV PATRIKES EGGRAFES
        '        If idPART = -1 Then  ' ΑΠΟ ΠΡΟΦΟΡΜΕΣ
        '            GDB.Execute("UPDATE PARTIDES SET YPOL=YPOL+" + poso + " WHERE ID=" + idT)
        '        Else
        '            GDB.Execute("UPDATE TIMS SET YPOL=YPOL+" + poso + " WHERE ID=" + idT)
        '        End If


        '        'ΕΝΗΜΕΡΩΝΩ ΠΑΡΑΧΘΕΙΣΕΣ ΠΑΡΤΙΔΕΣ
        '        GDB.Execute("UPDATE PARTIDES SET YPOL=YPOL-" + poso + " WHERE PARTIDA=" + mpART)
        '        GDB.Execute("DELETE FROM TIMSANAL WHERE ID=" + id)

        '    Next
        '    GDB.CommitTrans()

        'Catch ex As Exception

        '    GDB.RollbackTrans()
        '    Exit Sub

        'End Try

        'MsgBox("ΔΙΕΓΡΆΦΗ")
        'Me.Close()








    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        FrmAddEidos.n1.Text = n1.Text
        FrmAddEidos.IsNew = False

        '        Dim mpART As String = GridView1.CurrentRow.Cells("N2").Value.ToString
        '       Dim mORA As String = GridView1.CurrentRow.Cells("ΩΡΑ ΠΑΡΑΓ").Value.ToString

        FrmAddEidos.KOD.Text = GridView1.CurrentRow.Cells(1).Value.ToString
        FrmAddEidos.ONO.Text = GridView1.CurrentRow.Cells(0).Value.ToString
        FrmAddEidos.MON.Text = GridView1.CurrentRow.Cells(4).Value.ToString
        FrmAddEidos.BAROS.Text = GridView1.CurrentRow.Cells(3).Value.ToString
        FrmAddEidos.ID = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString
        'FrmAddSupplier.ID = GridView1.CurrentRow.Cells("BAROS").Value.ToString


        FrmAddEidos.ShowDialog()
        paint_ergasies()

    End Sub

    Private Sub analtimpol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles analtimpol.Click
        'Dim per = "Λίστα Τιμολογίων Πώλησης με Παρτίδες"
        'Dim kod = "00"
        'Dim matim As String = GridView1.CurrentRow.Cells(5).Value.ToString

        'Dim mergates As New ergates()
        'For k = 0 To 20
        '    mergates.widths(7) = 100
        'Next
        'Dim Mn1 As String
        'Mn1 = kod    '  Split(KATHG.Text, ";")(0)
        'mergates.Text = per '"Αρχείο Υλικών"
      



        Dim mergates As New ergates()
        'For k = 0 To 20
        '    mergates.widths(7) = 100
        'Next
        For KK As Integer = 0 To 6
            mergates.SUMES(KK) = 0
            mergates.widths(KK) = 100
        Next
        mergates.SUMES(4) = 1

        'Mn1 = GridView1(GridView1.CurrentRow)(GridView1.CurrentCell).tostring    '  Split(KATHG.Text, ";")(0)
        mergates.Text = "Τιμολόγια"
        'mergates.Label1.Text = "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2,ID  FROM YLIKA WHERE N1=" + Mn1 + " ORDER BY KOD "

        ' ο κωδικος του προιοντος που διαλεξα
        Dim matim As String = GridView1.CurrentRow.Cells(5).Value.ToString


        
        mergates.Label1.Text = "select TIMSPOL.KOD,ONO,HME,PARTIDA AS [ΠΑΡΤΙΔΑ],POSO AS [ΤΕΜΑΧ],ATIM as [ΤΙΜΟΛ.],TIMSPOL.ID  FROM TIMSPOL  INNER JOIN YLIKA ON TIMSPOL.KOD=YLIKA.KOD  where ATIM='" + matim + "' ORDER BY ID "

        ' mergates.Label1.Text = "SELECT T.N2 ,T.HME ,T.POSO,T.ID  FROM TIMSANAL T LEFT JOIN PARTIDES P ON T.IDPART=P.ID where IDTIMS='" + mk + "'"

        'mergates.Label1.Text = "SELECT KOD AS [ΚΩΔ],KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],ID  FROM SYNTAGES  where KOD='" + mk + "'"

        ' ergates.MdiParent = Me
        mergates.WindowState = FormWindowState.Maximized
        mergates.STHLHONOMATOS_ID = 0
        mergates.STHLHTOY_ID = 6
        mergates.widths(1) = 100
        mergates.QUERY_AFTER = ""  '"update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            mergates.widths(KK) = 100
        Next
        mergates.Label2.Text = "synt"  'per '"υλικα...."   ' KATHG.Text
        mergates.widths(1) = 200
        gMenu = 22
        mergates.Alignments(4) = DataGridViewContentAlignment.MiddleRight

        mergates.Read_Only = True
        mergates.ShowDialog()





    End Sub

    Private Sub symorfosi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditDomatia.Click

        Dim M_ID As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString()


        Dim F As New Fdomatia
        F.Text = GridView1.CurrentRow.Cells(1).Value.ToString()
        F.ID = M_ID
        F.HotelName.Text = GridView1.CurrentRow.Cells(0).Value.ToString()
        F.ShowDialog()






    End Sub
  






    Private Sub Button3_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dior_pel.Click
        If Me.Text = "Ξενοδοχεία" Then
            Dim M_ID As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString()
            Dim frmPEL_ITEM As New FrmADDSUPPLIER_ITEM
            frmPEL_ITEM.IsNew = False
            frmPEL_ITEM.ID = M_ID
            frmPEL_ITEM.N1.Text = M_ID
            frmPEL_ITEM.txtDie.Text = GridView1.CurrentRow.Cells(4).Value.ToString
            ' "select NAME,CATEGORY,EMAIL,THL,DIE,ID  FROM HOTELS "txtcategory
            frmPEL_ITEM.txtTHL.Text = GridView1.CurrentRow.Cells(3).Value.ToString
            frmPEL_ITEM.txtEmail.Text = GridView1.CurrentRow.Cells(2).Value.ToString
            frmPEL_ITEM.txtName.Text = GridView1.CurrentRow.Cells(0).Value.ToString


            frmPEL_ITEM.RANK.Text = GridView1.CurrentRow.Cells("RANK").Value.ToString
            frmPEL_ITEM.ComboCAT.Text = GridView1.CurrentRow.Cells("CATEGORY").Value.ToString

            'frmPEL_ITEM.txtcategory.Text = GridView1.CurrentRow.Cells(1).Value.ToString




            frmPEL_ITEM.ShowDialog()
        Else

            Dim frmPEL As New FrmAddSupplier

            'frmPEL.n1.Text = n1.Text
            frmPEL.IsNew = False

            '        Dim mpART As String = GridView1.CurrentRow.Cells("N2").Value.ToString
            '       Dim mORA As String = GridView1.CurrentRow.Cells("ΩΡΑ ΠΑΡΑΓ").Value.ToString


            frmPEL.cAirport.Text = GridView1.CurrentRow.Cells("AIRPORT").Value.ToString
            frmPEL.CH1.Text = GridView1.CurrentRow.Cells("CH1").Value.ToString
            frmPEL.CH2.Text = GridView1.CurrentRow.Cells("CH2").Value.ToString
            frmPEL.PTHSHC5.Text = GridView1.CurrentRow.Cells("CH5").Value.ToString

            frmPEL.PTHSHANAXC6.Text = GridView1.CurrentRow.Cells("CH6").Value.ToString

            frmPEL.rank.Text = GridView1.CurrentRow.Cells("RANK").Value.ToString

            frmPEL.KINHTO.Text = GridView1.CurrentRow.Cells("KINHTO").Value.ToString

            frmPEL.email.Text = GridView1.CurrentRow.Cells("EMAIL").Value.ToString
            frmPEL.ONO.Text = GridView1.CurrentRow.Cells("EPO").Value.ToString
            frmPEL.onoProsf.Text = GridView1.CurrentRow.Cells("ONO").Value.ToString
            frmPEL.Synodos.Text = GridView1.CurrentRow.Cells("SYNODOS").Value.ToString
            frmPEL.DIE.Text = GridView1.CurrentRow.Cells("DIE").Value.ToString
            frmPEL.ID = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString
            'frmPEL.ID = GridView1.CurrentRow.Cells("BAROS").Value.ToString  
            Dim C3 As String
            C3 = GridView1.CurrentRow.Cells("CH3").Value.ToString
            For L3 As Integer = 1 To 7
                If Mid(C3, L3, 1) = "1" Then
                    frmPEL.CheckedListBox1.SetItemChecked(L3 - 1, True)
                Else
                    frmPEL.CheckedListBox1.SetItemChecked(L3 - 1, False)
                End If
            Next
            Dim C4 As String
            C4 = GridView1.CurrentRow.Cells("CH4").Value.ToString
            For L3 = 1 To 16
                If Mid(C4, L3, 1) = "1" Then
                    frmPEL.CheckedListBox2.SetItemChecked(L3 - 1, True)
                Else
                    frmPEL.CheckedListBox2.SetItemChecked(L3 - 1, False)
                End If
            Next

            Try
                frmPEL.DTCheckin.Value = GridView1.CurrentRow.Cells("CHECKIN").Value.ToString
                frmPEL.DTCheckout.Value = GridView1.CurrentRow.Cells("CHECKOUT").Value.ToString


                frmPEL.DtAirAfixi.Value = GridView1.CurrentRow.Cells("AIRAFIXI").Value.ToString
                frmPEL.dtAirAnax.Value = GridView1.CurrentRow.Cells("AIRANAX").Value.ToString
                frmPEL.txtTo.Visible = True

            Catch ex As Exception

            End Try
            frmPEL.ShowDialog()
        End If

        paint_ergasies()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles add_pel.Click

        If Me.Text = "Ξενοδοχεία" Then
            'Dim M_ID As String = GridView1.CurrentRow.Cells(STHLHTOY_ID).Value.ToString()


            FrmADDSUPPLIER_ITEM.txtDie.Text = ""
            ' "select NAME,CATEGORY,EMAIL,THL,DIE,ID  FROM HOTELS "txtcategory
            FrmADDSUPPLIER_ITEM.txtTHL.Text = ""
            FrmADDSUPPLIER_ITEM.txtEmail.Text = ""
            FrmADDSUPPLIER_ITEM.txtName.Text = ""
            ' FrmADDSUPPLIER_ITEM.txtcategory.Text = ""


            FrmADDSUPPLIER_ITEM.IsNew = True

            FrmADDSUPPLIER_ITEM.ShowDialog()

            ' FrmADDSUPPLIER_ITEM.ID = M_ID
            ' FrmADDSUPPLIER_ITEM.N1.Text = M_ID

        Else


            FrmAddSupplier.email.Text = ""  'GridView1.CurrentRow.Cells(0).Value.ToString
            FrmAddSupplier.ONO.Text = ""  'GridView1.CurrentRow.Cells(1).Value.ToString
            FrmAddSupplier.onoProsf.Text = "" ' GridView1.CurrentRow.Cells(2).Value.ToString
            FrmAddSupplier.DIE.Text = ""  'GridView1.CurrentRow.Cells(3).Value.ToString




            ' FrmAddSupplier.n1.Text = n1.Text
            FrmAddSupplier.IsNew = True
            FrmAddSupplier.ShowDialog()
        End If

        paint_ergasies()

    End Sub
End Class