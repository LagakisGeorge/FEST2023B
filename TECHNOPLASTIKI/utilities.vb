Imports System.Data.OleDb
Imports System.Data.SqlClient

Imports System
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Windows.Forms
Imports System.Drawing.Printing
Imports System.Drawing
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6







Public Class utilities

    Dim GDB As New ADODB.Connection


    'Create connection
    Dim conn As OleDbConnection

    'create data adapter
    Dim da As OleDbDataAdapter

    'create dataset
    Dim ds As DataSet = New DataSet

    Dim dt As New DataTable



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '================================ PRODUCT_ATTRIBUTES ================================
        ExecuteSQLQuery("SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'TABLEFIELDS'")

        'On Error Resume Next

        If sqlDT(0)(0) = 0 Then
            ExecuteSQLQuery("CREATE TABLE TABLEFIELDS ( " & _
                            "ID INT IDENTITY(1,1)" & _
                            ",TABLENAME VARCHAR(16) NOT NULL" & _
                            ",FIELDNAME VARCHAR(30) NULL" & _
                            ",FIELDDESCRIPTION VARCHAR(50) NULL" & _
                            ",SUMES BIT  NULL" & _
                            ",FIELDTYPE VARCHAR(10) NULL " & _
                            ",COMBOQUERY VARCHAR(100) NULL )")

        End If


        Dim ANS As Integer = MsgBox("ΠΡΟΣΟΧΗ ΘΑ ΣΒΗΣΤΟΥΝ ΤΑ ΠΑΛΙΑ . EISAI SIGOYROS NAI /OXI", MsgBoxStyle.YesNo)

        If ANS = vbYes Then
            ExecuteSQLQuery("DELETE FROM TABLEFIELDS")
            UPDATE_TABLE("PEL")
            UPDATE_TABLE("EID")
            UPDATE_TABLE("TIM")
            UPDATE_TABLE("GRA")
            UPDATE_TABLE("EGG")
            UPDATE_TABLE("EGGTIM")
        End If

    End Sub

    Sub UPDATE_TABLE(ByVal PINAKAS As String)
        Dim DT As New DataTable
        ExecuteSQLQuery("SELECT TOP 1 * FROM " + PINAKAS, DT)
        Dim K As Integer
        For K = 0 To DT.Columns.Count - 1
            ExecuteSQLQuery("INSERT INTO TABLEFIELDS (TABLENAME,FIELDNAME) VALUES" & _
                             "( '" + PINAKAS + "','" + DT.Columns(K).Caption + "')")
        Next
        DT = Nothing

    End Sub





    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ANS As Integer = MsgBox("ΠΡΟΣΟΧΗ ΘΑ ΣΒΗΣΤΟΥΝ ΤΑ ΠΑΛΙΑ . EIΣAI ΣΙΓΟΥΡΟΣ NAI /OXI", MsgBoxStyle.YesNo)

        If ANS = vbNo Then Exit Sub


        '================================ PRODUCT_ATTRIBUTES ================================
        ExecuteSQLQuery("SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'NEWMENUS'")

        'On Error Resume Next




        If sqlDT(0)(0) = 1 Then
            ExecuteSQLQuery("DROP TABLE NEWMENUS")
        End If

        ExecuteSQLQuery("CREATE TABLE NEWMENUS ( " & _
                        "ID INT IDENTITY(1,1)" & _
                        ",[MENU_ID] INT NULL " & _
       ",[PARENT_ID] INT NULL " & _
       ",[SQL1] TEXT NULL " & _
       ",[SQL2] TEXT NULL " & _
        ",[TSQLSELECT] Text NULL " & _
         ",[TSQLWHERE] Text NULL " & _
          ",[TSQLORDER] Text NULL " & _
       ",[MENUNAME] VARCHAR(200) NULL) ")






    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub


    Private Sub SQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sql.Click

        ' a tropos   ===========================================
        'Try
        '    paint_ergasies(TextBox1.Text)
        'Catch ex As Exception

        'End Try



        'a2 tropos

        'Dim da As New OleDbDataAdapter
        'DataGridView1.ClearSelection()
        'DataGridView1.Columns.Clear()
        '' ExecuteSQLQuery(SQLqry, mds)
        'ExecuteSQLQuery(TextBox1.Text, dt, da)

        'DataGridView1.DataSource = ds
        'DataGridView1.Refresh()
        'mds = Nothing





        ' b tropos  ============================================

        Dim conn As New OleDbConnection
        conn.ConnectionString = gConnect
        conn.Open()



        Try

            da = New OleDbDataAdapter(TextBox1.Text, conn)

            'create command builder
            ' Dim cb As OleDbCommandBuilder = New OleDbCommandBuilder(da)
            ds.Clear()
            'fill dataset
            'Exit Sub
            Try
                da.Fill(ds, "PEL")
                DataGridView1.ClearSelection()
                DataGridView1.DataSource = ds
                DataGridView1.DataMember = "PEL"



            Catch ex As Exception
                MsgBox(Err.Description + Chr(13) + TextBox1.Text)

            End Try

            ' Exit Sub

            'GridView1.Columns(STHLHTOY_ID).Width = 0
            ' DataGridView1.Columns(STHLHTOY_ID).Visible = False

        Catch ex As SqlException
            MsgBox(ex.ToString)
        Finally
            ' Close connection
            conn.Close()
        End Try




    End Sub

    Public Sub paint_ergasies(ByVal sql As String)

        Dim mds As New DataTable
        Dim cnString As String = gConSQL
        Dim SQLqry As String = sql '"SELECT NAME,N1,ID FROM ERGATES " ' ORDER BY HME "
        Try
            DataGridView1.ClearSelection()
            DataGridView1.Columns.Clear()
            ExecuteSQLQuery(SQLqry, mds)
            DataGridView1.DataSource = mds
            DataGridView1.Refresh()
            mds = Nothing

        Catch ex As SqlException
            MsgBox(ex.ToString)
        Finally
            ' Close connection
            'MDS = Nothing
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click



        'R.Open(TextBox1.Text, GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        'Me.Text = R.Fields("EPO").Value
        'R.Close()
        Dim K As Integer

        Try
            GDB.Execute(TextBox1.Text, K)
        Catch ex As Exception

        End Try


        MsgBox(K)

        Exit Sub




        'Try
        'da.Update(ds, "PEL")
        'Catch ex As Exception
        ' MsgBox("δεν αποθηκευτηκε" + ex.Message)
        'End Try
        ExecuteSQLQuery(TextBox1.Text)
        ' MsgBox(sqlDT.Rows.Count)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ''ListBox1.c()

        'ListBox1.FormattingEnabled = True
        'ListBox1.HorizontalScrollbar = True
        'Dim k As Integer
        'For k = 1 To 10

        '    ListBox1.Items.Add(k & "bbbb" & k & "ccxcx")
        'Next

        '' ListBox1.Items.AddRange(New Object() {"Item 1, column 1", "Item 2, column 1", "Item 3, column 1", "Item 4, column 1", "Item 5, column 1", "Item 1, column 2", "Item 2, column 2", "Item 3, column 2"})
        ''ListBox1.Location = New System.Drawing.Point(0, 0)
        'ListBox1.MultiColumn = True
        'ListBox1.Name = "listBox1"
        'ListBox1.ScrollAlwaysVisible = True
        '' ListBox1.Size = New System.Drawing.Size(120, 95)
        'ListBox1.TabIndex = 0
        'ListBox1.ColumnWidth = 85

    End Sub

    Private Sub utilities_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'FocusedCombo1.EnterFocusColor = Color.Yellow
        ' FocusedTextBox1.EnterFocusColor = Color.Yellow
        
        '  Private Sub cmdCommand1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCommand1.Click

        Dim R As New ADODB.Recordset

        GDB.Open(gConnect)
    End Sub

    Private Sub OPENFILE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPENFILE.Click
        OpenFileDialog1.ShowDialog()
        If Len(OpenFileDialog1.FileName) = 0 Then
            Exit Sub
        End If
        'If "OpenFileDialog1" Then
        Dim line As String
        Dim line2 As String
        Using sr As StreamReader = New StreamReader(OpenFileDialog1.FileName, System.Text.Encoding.Default)
            line = sr.ReadLine()
            'f_mess_pel = sr.ReadLine()
            'f_mess_eid = sr.ReadLine()
            'F_PLATH_PEL = sr.ReadLine()
            'F_PLATH_EID = sr.ReadLine()

            'εδω εχω τις ιδιαιτερότητες του πελάτη
            line2 = sr.ReadLine()   ' 5Η  SEIRA
        End Using

        TextBox1.Text = line


    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim SW As StreamWriter

        SaveFileDialog1.ShowDialog()
        Dim F As String
        F = SaveFileDialog1.FileName
        If Not File.Exists(F) Then
            SW = New StreamWriter(F, False, System.Text.Encoding.Default)

            ' sw.WriteLine("  ")
        Else
            SW = File.CreateText(F)
        End If

        SW.WriteLine(TextBox1.Text)

        SW.Close()


    End Sub





    Private Sub ΠελατώνToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ΠελατώνToolStripMenuItem.Click


        Dim ANS = MsgBox("Είναι η Γραμογράφηση ΠΕΛΑΤΩΝ EXCEL : ΚΩΔ;ΕΠΩΝΥΜΙΑ;ΔΙΕΥΘΥΝΣΗ;ΑΦΜ;ΤΗΛ¨ ??", MsgBoxStyle.YesNo)
        If ANS = MsgBoxResult.No Then
            Exit Sub
        End If

        Me.Text = "Πελατών"

        OpenFileDialog1.ShowDialog()

        Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim c As String



        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value


        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)

        Dim N As Integer = 1


        Do

            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString
            If xlWorkSheet.Cells(N, 1).VALUE = Nothing Then

                Exit Do
            End If

            Try

                c = xlWorkSheet.Cells(N, 1).VALUE.ToString
                ' c = "000000" + c

                r.Open("select * FROM PEL WHERE EIDOS='e' AND KOD='" + c + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If r.EOF Then
                    GDB.Execute("INSERT INTO PEL (EIDOS,KOD) VALUES ('e','" + c + "')")
                End If
                r.Close()
                GDB.Execute("update PEL SET EPO='" + xlWorkSheet.Cells(N, 2).VALUE.ToString + "',DIE='" + xlWorkSheet.Cells(N, 3).VALUE.ToString + "',AFM='" + xlWorkSheet.Cells(N, 4).VALUE.ToString + "',THL='" + xlWorkSheet.Cells(N, 5).VALUE.ToString + "' WHERE EIDOS='e' and KOD='" + c + "'")
            Catch ex As Exception
                MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + c + " == " + ex.Message)
                'Exit Do

            End Try


            N = N + 1
            Me.Text = N
        Loop Until xlWorkSheet.Cells(N, 1).VALUE Is Nothing



        MsgBox("OK")
















        'xlWorkBook.Save()

        'xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        'xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)



        Exit Sub

































        'Using sr As StreamReader = New StreamReader(OpenFileDialog1.FileName, System.Text.Encoding.Default)


        '    Do

        '        line = sr.ReadLine()
        '        If line Is Nothing Then
        '            Exit Do
        '        End If

        '        Try

        '            c = Split(line, ";")(0)
        '            c = "000000" + c

        '            r.Open("select * FROM PEL WHERE EIDOS='e' AND KOD='" + c + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        '            If r.EOF Then
        '                GDB.Execute("INSERT INTO PEL (EIDOS,KOD) VALUES ('e','" + c + "')")
        '            End If
        '            r.Close()
        '            GDB.Execute("update PEL SET EPO='" + Split(line, ";")(1) + "',DIE='" + Split(line, ";")(2) + "',AFM='" + Split(line, ";")(3) + "',THL='" + Split(line, ";")(4) + "' WHERE EIDOS='e' and KOD='" + c + "'")
        '        Catch ex As Exception
        '            MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + c + " == " + ex.Message)
        '            'Exit Do

        '        End Try




        '    Loop Until line Is Nothing




        'End Using



    End Sub

    Private Sub ΠρομηθευτώνToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ΠρομηθευτώνToolStripMenuItem.Click
        Dim ANS = MsgBox("Είναι η Γραμογράφηση EXCEL ΠΡΟΜΗΘΕΥΤΩΝ : ΚΩΔ;ΕΠΩΝΥΜΙΑ;ΔΙΕΥΘΥΝΣΗ;ΑΦΜ;ΤΗΛ¨ ??", MsgBoxStyle.YesNo)
        If ANS = MsgBoxResult.No Then
            Exit Sub
        End If



        OpenFileDialog1.ShowDialog()

        Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim c As String


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value


        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)

        Dim N As Integer = 1


        Do

            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString
            If xlWorkSheet.Cells(N, 1).VALUE = Nothing Then

                Exit Do
            End If

            Try

                c = xlWorkSheet.Cells(N, 1).VALUE.ToString
                ' c = "000000" + c

                r.Open("select * FROM PEL WHERE EIDOS='r' AND KOD='" + c + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If r.EOF Then
                    GDB.Execute("INSERT INTO PEL (EIDOS,KOD) VALUES ('r','" + c + "')")
                End If
                r.Close()
                GDB.Execute("update PEL SET EPO='" + xlWorkSheet.Cells(N, 2).VALUE.ToString + "',DIE='" + xlWorkSheet.Cells(N, 3).VALUE.ToString + "',AFM='" + xlWorkSheet.Cells(N, 4).VALUE.ToString + "',THL='" + xlWorkSheet.Cells(N, 5).VALUE.ToString + "' WHERE EIDOS='r' and KOD='" + c + "'")
            Catch ex As Exception
                MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + c + " == " + ex.Message)
                'Exit Do

            End Try


            N = N + 1
            Me.Text = N
        Loop Until xlWorkSheet.Cells(N, 1).VALUE Is Nothing



        MsgBox("OK")
















        'xlWorkBook.Save()

        'xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        'xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
















        Using sr As StreamReader = New StreamReader(OpenFileDialog1.FileName, System.Text.Encoding.Default)


            Do

                line = sr.ReadLine()
                If line Is Nothing Then
                    Exit Do
                End If

                Try

                    c = Split(line, ";")(0)

                    r.Open("select * FROM PEL WHERE EIDOS='r' AND KOD='" + c + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    If r.EOF Then
                        GDB.Execute("INSERT INTO PEL (EIDOS,KOD) VALUES ('r','" + c + "')")
                    End If
                    r.Close()

                    GDB.Execute("update PEL SET EPO='" + Split(line, ";")(1) + "',DIE='" + Split(line, ";")(2) + "',AFM='" + Split(line, ";")(3) + "',THL='" + Split(line, ";")(4) + "' WHERE EIDOS='r' and KOD='" + c + "'")
                Catch ex As Exception
                    MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + c + Chr(13) + ex.Message)
                    Exit Do

                End Try




            Loop Until line Is Nothing




        End Using

    End Sub

    Private Sub ΕιδώνToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ΕιδώνToolStripMenuItem.Click
        Dim ANS = MsgBox("Είναι η Γραμογράφηση EXCEL : ΚΩΔ;ONOMA;ΚΑΤΗΓΟΡΙΑ 1-5;μοναδα ??" + Chr(13) + "Αν θέλετε να ενημερώσετε μονο μία στήλη,βαζετε κωδικο & κενές στηλες αυτές που θα μείνουν ίδιες", MsgBoxStyle.YesNo)
        If ANS = MsgBoxResult.No Then
            Exit Sub
        End If



        OpenFileDialog1.ShowDialog()

        Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim c As String


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value


        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)

        Dim N As Integer = 1


        Do

            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString
            If xlWorkSheet.Cells(N, 1).VALUE = Nothing Then

                Exit Do
            End If

            Try

                c = xlWorkSheet.Cells(N, 1).VALUE.ToString
                ' c = "000000" + c

                r.Open("select * FROM YLIKA WHERE  KOD='" + c + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If r.EOF Then
                    GDB.Execute("INSERT INTO YLIKA (KOD) VALUES ('" + c + "')")
                End If
                r.Close()


                If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 2).VALUE) Then
                Else
                    GDB.Execute("update YLIKA SET ONO='" + Replace(xlWorkSheet.Cells(N, 2).VALUE.ToString, "'", "`") + "' WHERE  KOD='" + c + "'")
                End If



                If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 3).VALUE) Then
                Else
                    GDB.Execute("update YLIKA SET N1=" + xlWorkSheet.Cells(N, 3).VALUE.ToString + " WHERE  KOD='" + c + "'")
                End If

                If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 4).VALUE) Then
                Else
                    GDB.Execute("update YLIKA SET C1='" + Replace(xlWorkSheet.Cells(N, 4).VALUE.ToString, "'", "`") + "' WHERE  KOD='" + c + "'")
                End If



                '  GDB.Execute("update YLIKA SET ONO='" + Replace(xlWorkSheet.Cells(N, 2).VALUE.ToString, "'", "`") + "',N1=" + xlWorkSheet.Cells(N, 3).VALUE.ToString + " WHERE  KOD='" + c + "'")





            Catch ex As Exception
                MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + c + " == " + ex.Message)
                'Exit Do

            End Try


            N = N + 1
            Me.Text = N
        Loop Until xlWorkSheet.Cells(N, 1).VALUE Is Nothing



        MsgBox("OK")
















        'xlWorkBook.Save()

        'xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        'xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)













    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Dim ANS = MsgBox("Είναι η Γραμογράφηση EXCEL : Β-ΚΩΔ;D-ONOMA;H-βαρος ??", MsgBoxStyle.YesNo)
        If ANS = MsgBoxResult.No Then
            Exit Sub
        End If



        OpenFileDialog1.ShowDialog()

        Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim c As String


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value


        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)

        Dim N As Integer = 1


        Do

            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString
            If xlWorkSheet.Cells(N, 2).VALUE = Nothing Then

                Exit Do
            End If

            Try

                c = xlWorkSheet.Cells(N, 2).VALUE.ToString
                ' c = "000000" + c

                r.Open("select * FROM YLIKA WHERE  KOD='" + c + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If r.EOF Then
                    GDB.Execute("INSERT INTO YLIKA (KOD,N1) VALUES ('" + c + "',4)")
                End If
                r.Close()
                GDB.Execute("update YLIKA SET ONO='" + Replace(xlWorkSheet.Cells(N, 4).VALUE.ToString, "'", "`") + "',BAROS=" + xlWorkSheet.Cells(N, 8).VALUE.ToString + " WHERE  KOD='" + c + "'")
            Catch ex As Exception
                MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + Str(N) + " " + c + " == " + ex.Message)
                'Exit Do

            End Try


            N = N + 1
            Me.Text = N
        Loop Until xlWorkSheet.Cells(N, 2).VALUE Is Nothing



        MsgBox("OK " + Str(N))





    End Sub

    Private Sub ΣυνταγωνToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ΣυνταγωνToolStripMenuItem.Click
        Dim ANS = MsgBox("Είναι η Γραμογράφηση EXCEL : I-ΚΩΔ ΠΡΟΙΟΝ;J-ΚΩΔ.ΣΥΣΤΑ;K-ΠΕΡ%  ??", MsgBoxStyle.YesNo)
        If ANS = MsgBoxResult.No Then
            Exit Sub
        End If



        OpenFileDialog1.ShowDialog()

        Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim c As String


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value


        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)

        Dim N As Integer = 1
        Dim D As String


        Do

            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString
            If xlWorkSheet.Cells(N, 9).VALUE = Nothing Then

                Exit Do
            End If

            Try

                c = xlWorkSheet.Cells(N, 9).VALUE.ToString
                D = xlWorkSheet.Cells(N, 10).VALUE.ToString

                ' c = "000000" + c

                r.Open("select * FROM SYNTAGES WHERE KODSYNOD='" + D + "' AND  KOD='" + c + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If r.EOF Then
                    GDB.Execute("INSERT INTO SYNTAGES (KOD,KODSYNOD) VALUES ('" + c + "','" + D + "')")
                End If
                r.Close()
                GDB.Execute("update SYNTAGES SET POSOSTO=" + Replace(Str(xlWorkSheet.Cells(N, 11).VALUE / 1), ",", ".") + " WHERE  KOD='" + c + "' AND KODSYNOD='" + D + "'")
            Catch ex As Exception
                MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + Str(N) + " " + c + " == " + D + "--" + ex.Message)
                'Exit Do

            End Try


            N = N + 1
            Me.Text = N
        Loop Until xlWorkSheet.Cells(N, 9).VALUE Is Nothing



        MsgBox("OK " + Str(N))



    End Sub

    Private Sub ΤιμολόγιαΑγοράςToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ΤιμολόγιαΑγοράςToolStripMenuItem.Click

      



    End Sub

    Private Sub ΤιμολόγιαΠώλησηςToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ΤιμολόγιαΠώλησηςToolStripMenuItem.Click
        Dim Printer As New Printer
        Printer.Print("Total (" & (19 + 300) / 4 & ")")
        Printer.EndDoc()



    End Sub



   
    Private Sub SQL_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SQL.Click

        'a2 tropos

        'Dim da As New OleDbDataAdapter
        'DataGridView1.ClearSelection()
        'DataGridView1.Columns.Clear()
        '' ExecuteSQLQuery(SQLqry, mds)
        'ExecuteSQLQuery(TextBox1.Text, dt, da)

        'DataGridView1.DataSource = ds
        'DataGridView1.Refresh()
        'mds = Nothing





        ' b tropos  ============================================

        Dim conn As New OleDbConnection
        conn.ConnectionString = gConnect
        conn.Open()



        Try

            da = New OleDbDataAdapter(TextBox1.Text, conn)

            'create command builder
            ' Dim cb As OleDbCommandBuilder = New OleDbCommandBuilder(da)
            ds.Clear()
            'fill dataset
            'Exit Sub
            Try
                da.Fill(ds, "PEL")
                DataGridView1.ClearSelection()
                DataGridView1.DataSource = ds
                DataGridView1.DataMember = "PEL"



            Catch ex As Exception
                MsgBox(Err.Description + Chr(13) + TextBox1.Text)

            End Try

            ' Exit Sub

            'GridView1.Columns(STHLHTOY_ID).Width = 0
            ' DataGridView1.Columns(STHLHTOY_ID).Visible = False

        Catch ex As SqlException
            MsgBox(ex.ToString)
        Finally
            ' Close connection
            conn.Close()
        End Try



    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim LIST As New List(Of String)()
        LIST.Add("1 AASSSSS")
        LIST.Add("2 AASSSSS")
        LIST.Add("3 AASSSSS")
        Me.Text = LIST.Item(1).ToString
        ' ΠΟΣΑ ΜΕΛΗ ΕΧΕΙ Ο ΠΙΝΑΚΑΣ :  LIST.COUNT

    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim C As String = "FESTIVAL.BAK"

100:    C = InputBox("ΔΩΣΕ ΟΝΟΜΑ BACKUP ΤΗΣ ΒΑΣΗΣ THERMO ΣΤΟ FEST.bak", , C)
110:    ExecuteSQLQuery("BACKUP DATABASE [TECHNOPLASTIKI] TO  DISK ='" + C + "' WITH NOFORMAT, NOINIT, SKIP, NOREWIND, NOUNLOAD,  STATS = 10")
120:    MsgBox("ΟΛΟΚΛΗΡΩΘΗΚΕ")

    End Sub

    Private Sub ΔΙΑΓΡΑΦΗΠΑΡΤΙΔΑΣΜΕIDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ΔΙΑΓΡΑΦΗΠΑΡΤΙΔΑΣΜΕIDToolStripMenuItem.Click
        Dim M_ID As String
        M_ID = InputBox("ΔΩΣΕ ΤΟ  ID ")

        Dim PARTIDES As New DataTable

        ExecuteSQLQuery("SELECT * FROM PARTIDES WHERE ID=" + M_ID, partides)
        If PARTIDES.Rows.Count = 0 Then
            MsgBox(" ΔΕΝ ΥΠΑΡΧΕΙ ΕΓΓΡΑΦΗ ΜΕ ΑΥΤΟ ΤΟ ID")
            Exit Sub

        Else
            Dim N As Integer = MsgBox("ΝΑ ΔΙΑΓΡΑΦΕΙ Η ΠΑΡΤΙΔΑ " + PARTIDES.Rows(0)("PARTIDA").ToString(), MsgBoxStyle.YesNo)
            If N = vbYes Then
                ExecuteSQLQuery("DELETE FROM PARTIDES WHERE ID=" + M_ID, PARTIDES)
                MsgBox("ΟΚ ΔΙΑΓΡΑΦΗΚΕ")
            Else
                MsgBox("ΔΕΝ ΔΙΑΓΡΑΦΗΚΕ")
            End If








        End If




    End Sub
End Class