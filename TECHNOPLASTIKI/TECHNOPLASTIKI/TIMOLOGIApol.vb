
Imports Excel = Microsoft.Office.Interop.Excel


Imports System.Data.OleDb
Imports System.Net.NetworkInformation



Public Class TIMOLOGIApol
    Dim GDB As New ADODB.Connection

    '  Private Sub cmdCommand1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCommand1.Click

    Dim R As New ADODB.Recordset
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' ExecuteSQLQuery("SELECT * FROM EID WHERE ONO LIKE '" + EIDOS.Text + "%'")
        If Len(EIDKOD.Text) = 0 Then
            FillListBox("SELECT ONO,KOD   FROM YLIKA WHERE   ONO LIKE '" + EIDOS.Text + "%'", ListBox1)
        Else
            FillListBox("SELECT ONO,KOD   FROM YLIKA WHERE  KOD LIKE '" + EIDKOD.Text + "%'", ListBox1)
        End If





    End Sub



    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim N As Long

        'Dim i As Integer
        'Dim str() As String

        'str() = Split("0;1;2;3;4;5", ";")

        'For i = 0 To 5
        '    List1.AddItem(str(i), i)
        'Next i




        N = ListBox1.SelectedIndex
        ONO.Text = Split(ListBox1.Items(N).ToString, ";")(0)
        KOD.Text = Split(ListBox1.Items(N).ToString, ";")(1)



        'FillListBox("SELECT left(ATIM+space(10),10)+CONVERT(CHAR(10),HME,3),str(YPOL,6,1)+' '+PROM    FROM TIMS WHERE YPOL>0  AND KOD='" + KOD.Text + "' ORDER BY HME", ListBox2)
        'ListBox2.Items.Insert(0, "ΠΑΡ/ΚΟ    HMEΡ/ΝΙΑ      ΥΠΟΛ. ΠΡΟΜΗΘΕΥΤΗΣ ")

        katax.Enabled = True





        HME.Focus()
        ' ExecuteSQLQuery("UPDATE MEM SET MEMO='" + Split(ListBox1.Items(N).ToString, ";")(0) + "' WHERE ID=1")
    End Sub

    Private Sub katax_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles katax.Click
        Dim sql As String
        If Val(POSO.Text) = 0 Then
            MsgBox("Ποσότητα;")
            Exit Sub
        End If
        If Len(ATIM.Text) = 0 Then
            MsgBox("Αρ.Τιμολογίου;")
            Exit Sub
        End If


        If Len(pel.Text) = 0 Then
            MsgBox("Πελάτης;")
            Exit Sub
        End If




        Dim mposo As String
        Dim mgr As Double
        Dim MPART As String
        Dim temPart As Long



        Dim MFI As New DataTable
        ExecuteSQLQuery("select count(*) from TIMSPOL WHERE ATIM='" + ATIM.Text + "' AND HME='" + Format(HME.Value, "MM/dd/yyyy") + "'", MFI)

        If MFI(0)(0) > 0 Then
            MsgBox("ΥΠΑΡΧΕΙ ΗΔΗ ΤΟ ΤΙΜΟΛΟΓΙΟ")
            Exit Sub

        End If





        Dim idtimols As String
        Dim UPDATED As Boolean = False


        Dim ypoloipo As Long = Val(POSO.Text)
        Dim IDPART As String


        For K = 0 To PARTIDES.SelectedItems.Count - 1

            UPDATED = True

            MPART = Split(PARTIDES.SelectedItems(K).ToString(), ";")(0)
            temPart = Split(PARTIDES.SelectedItems(K).ToString(), ";")(2)
            IDPART = (Split(PARTIDES.SelectedItems(K).ToString(), ";")(2))

            If ypoloipo < temPart Then
                temPart = ypoloipo
            End If


            Dim FF As String


            If temPart > 0 And Len(pel.Text) > 0 And Len(KOD.Text) > 0 Then
                FF = Split(pel.Text, ";")(0)
                FF = Replace(FF, "'", "`")
                FF = Mid(FF, 1, 20)
                sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
                sql = sql + "'" + Format(HME.Value, "MM/dd/yyyy") + "',"
                sql = sql + Str(temPart) + ","
                sql = sql + "'" + ATIM.Text + "',"
                sql = sql + "'" + LTrim(KOD.Text) + "',"
                sql = sql + "'" + FF + "',"
                sql = sql + Str(MPART) + ",'" + Mid(Split(pel.Text, ";")(1), 1, 9) + "'," + IDPART + " )"
                Try
                    GDB.Execute(sql)
                    'ExecuteSQLQuery(sql)
                    katax.Enabled = False
                    ExecuteSQLQuery("select max(ID) FROM TIMSPOL")
                    idtimols = sqlDT.Rows(0)(0).ToString
                    GDB.Execute("UPDATE PARTIDES SET YPOL=YPOL-" + Str(temPart) + " WHERE PARTIDA=" + MPART)


                Catch ex As Exception
                    MsgBox("Δεν Αποθηκεύθηκε " + Err.Description)
                    Exit Sub
                End Try
                ypoloipo = ypoloipo - temPart


                LINES.Items.Add(KOD.Text + ";" + Str(MPART) + ";" + Str(temPart) + ";" + Space(100) + ";" + IDPART + ";" + idtimols)

            End If



        Next

        If UPDATED = False Then ' ΔΕΝ ΕΧΕΙ ΠΑΡΤΙΔΕΣ

            Dim RSYN As New ADODB.Recordset
            RSYN.Open("SELECT COUNT(*) FROM SYNTAGES WHERE KOD='" + LTrim(KOD.Text) + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If RSYN(0).Value > 0 Then  ' EINAI ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ



                'LINES.Items.Add(";;;;ΔΕΝ ΥΠΑΡΧΕΙ ΠΑΡΤΙΔΑ ΓΙΑ ΤΟ ΕΙΔΟΣ " + cEID)
            Else ' ΔΕΝ ΕΧΕΙ ΠΑΡΤΙΔΕΣ  //ΕΙΝΑΙ ΕΜΠΟΡΕΥΜΑ
                temPart = Val(POSO.Text)
                If UPDATE_TIMOL_AGORAS_AYLON() Then

                Else
                    Exit Sub
                End If
            End If


            Dim FF As String
            FF = Split(pel.Text, ";")(0)
            FF = Replace(FF, "'", "`")
            FF = Mid(FF, 1, 20)

            sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
            sql = sql + "'" + Format(HME.Value, "MM/dd/yyyy") + "',"
            sql = sql + Str(temPart) + ","
            sql = sql + "'" + ATIM.Text + "',"
            sql = sql + "'" + LTrim(KOD.Text) + "',"
            sql = sql + "'" + FF + "',"
            sql = sql + Str(0) + ",'" + Mid(Split(pel.Text, ";")(1).ToString, 1, 9) + "'," + "0" + " )"
            Try
                GDB.Execute(sql)
            Catch ex As Exception
                MsgBox(Err.Description + " " + Chr(13) + sql)


            End Try



        End If







        ' LINES.Items.Add("------------------------------------------")


        POSO.Text = ""

        PARTIDES.Items.Clear()


        Dim N As Long





        'N = ListBox1.SelectedIndex
        'ONO.Text = Split(ListBox1.Items(N).ToString, ";")(0)
        'KOD.Text = Split(ListBox1.Items(N).ToString, ";")(1)


        'FillListBox("SELECT left(ATIM+space(10),10)+CONVERT(CHAR(10),HME,3),str(YPOL,6,1)+' '+PROM    FROM TIMS WHERE YPOL>0  AND KOD='" + KOD.Text + "' ORDER BY HME", ListBox2)
        'ListBox2.Items.Insert(0, "ΠΑΡ/ΚΟ    HMEΡ/ΝΙΑ      ΥΠΟΛ. ΠΡΟΜΗΘΕΥΤΗΣ ")
        katax.Enabled = True










    End Sub

    Private Function UPDATE_TIMOL_AGORAS_AYLON() As Boolean
        UPDATE_TIMOL_AGORAS_AYLON = True
        Dim mFiFo As String
        If Mid(ComboFifo.Text, 1, 1) = 1 Then
            mFiFo = ""
        Else
            mFiFo = " desc"
        End If


        Dim mHME As String = Format(HME.Value, "MM/dd/yyyy") 'Mid(mHME, 4, 2) + "/" + Mid(mHME, 1, 2) + "/" + Mid(mHME, 7, 4)

        'Dim FF As String
        'FF = Split(pel.Text, ";")(0)
        'FF = Replace(FF, "'", "`")
        'FF = Mid(FF, 1, 20)

        'Sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
        'Sql = Sql + "'" + Format(HME.Value, "MM/dd/yyyy") + "',"
        'Sql = Sql + Str(temPart) + ","
        'Sql = Sql + "'" + ATIM.Text + "',"
        'Sql = Sql + "'" + LTrim(KOD.Text) + "',"
        'Sql = Sql + "'" + FF + "',"
        'Sql = Sql + Str(0) + ",'" + Mid(Split(pel.Text, ";")(1).ToString, 1, 9) + "'," + "0" + " )"











        'ΔΕΝ ΕΙΝΑΙ  ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ  ΑΡΑ ΕΙΝΑΙ ΠΡΩΤΗ ΥΛΗ Ή ΒΟΗΘΗΤΙΚΗ
        ' ΒΡΕΣ ΤΑ ΤΙΜΟΛΟΓΙΑ ΠΟΥ ΕΧΟΥΝ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ

        Dim isA_YLH As Integer = 1
        Dim CEID As String = LTrim(KOD.Text)
        Dim mATIM As String = ATIM.Text
        Dim mPOSO As Double = Val(POSO.Text)



        Dim POS2 As New DataTable

        Dim r2 As New ADODB.Recordset
        r2.Open("select sum(YPOL) FROM TIMS WHERE RTRIM(LTRIM(KOD))='" + CEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If IsDBNull(r2(0).Value) Then
            ' < mposo Then
            MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο TIM.AΓOΡAΣ για τον κωδικό " + CEID + Chr(13) + " το τιμολόγιο " + mATIM + " δεν θα περαστεί")
            UPDATE_TIMOL_AGORAS_AYLON = False
            Exit Function
        Else
            If r2(0).Value < mPOSO Then
                MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο TIM.AΓOΡAΣ για τον κωδικό " + CEID + Chr(13) + " το τιμολόγιο " + mATIM + " δεν θα περαστεί")

                UPDATE_TIMOL_AGORAS_AYLON = False
                Exit Function
            End If



        End If
        r2.Close()



        ExecuteSQLQuery("SELECT  KOD,YPOL,ATIM,HME,ID  FROM TIMS where YPOL>0 AND RTRIM(LTRIM(KOD))='" + CEID + "' ORDER BY HME " + mFiFo, POS2)
        If POS2.Rows.Count = 0 Then  ' '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Else '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim LL As Integer, m_MPOSO As Single = mPOSO
            '=======///==============
            For LL = 0 To POS2.Rows.Count - 1
                'ΤΟ ΠΟΣΟ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΕΙ ΕΙΝΑΙ ΛΙΓΟΤΕΡΟ ΑΠΟ ΤΟ ΤΙΜΟΛ.ΑΓΟΡΑΣ
                'ΟΠΟΤΕ ΑΦΑΙΡΩ ΟΛΟ ΤΟ ΠΟΣΟ ΤΟΥ ΤΙΜΟΛ.ΠΩΛΗΣΗΣ (M_POSO)
                If m_MPOSO < POS2.Rows(LL)("YPOL") Then
                    'DEBUG
                    'If check_only = False Then
                    GDB.Execute("UPDATE TIMS SET YPOL=YPOL-" + Replace(Str(m_MPOSO), ",", ".") + " WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                    GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(m_MPOSO), ",", ".") + ",'" + CEID + "','" + mATIM + "')")
                    'End If

                    m_MPOSO = 0

                Else '  m_MPOSO > POS2.Rows(LL)("YPOL") Then
                    'ΤΟ ΠΟΣΟ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΕΙ ΕΙΝΑΙ MEGALYTERO ΑΠΟ ΤΟ ΤΙΜΟΛ.ΑΓΟΡΑΣ
                    'ΟΠΟΤΕ ΑΦΑΙΡΩ ΟΛΟ ΤΟ ΠΟΣΟ ΤΟΥ ΤΙΜΟΛ.AGORAS POS2.Rows(LL)("YPOL")
                    m_MPOSO = m_MPOSO - POS2.Rows(LL)("YPOL")
                    'DEBUG
                    'If check_only = False Then
                    GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(POS2.Rows(LL)("YPOL")), ",", ".") + ",'" + CEID + "','" + mATIM + "')")
                    GDB.Execute("UPDATE TIMS SET YPOL=0 WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                    'End If

                End If
                If m_MPOSO = 0 Then
                    Exit For
                End If
            Next
            '========///=============
        End If  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    End Function





    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        For k = 0 To 20
            ergates.widths(7) = 100
        Next
        '  ExecuteSQLQuery("update YLIKA SET N1=(SELECT TOP 1 AEG FROM EID WHERE KOD=YLIKA.KOD) ")

        Dim Mn1 As String
        Mn1 = "" ' Split(KATHG.Text, ";")(0)


        ergates.Text = "Αρχείο Υλικών"
        ergates.Label1.Text = "SELECT HME,POSO,ATIM,KOD,PROM,YPOL,ID  FROM TIMSPOL ORDER BY HME "

        ' ergates.MdiParent = Me
        ergates.WindowState = FormWindowState.Maximized
        ergates.STHLHONOMATOS_ID = 6
        ergates.STHLHTOY_ID = 6
        ergates.widths(1) = 100
        ergates.QUERY_AFTER = " " 'update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            ergates.widths(KK) = 50
        Next
        ergates.Label2.Text = "Τιμολόγια"
        ergates.widths(0) = 100
        gMenu = 24
        ergates.Show()
    End Sub

    Private Sub TIMOLOGIApol_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave
        'Dim ans = MsgBox("Να αποθηκευτούν οι αλλαγές;", MsgBoxStyle.YesNo)
        'If ans = MsgBoxResult.Yes Then
        '    GDB.CommitTrans()
        'Else
        '    GDB.RollbackTrans()
        'End If
    End Sub

    Private Sub TIMOLOGIA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'monades.Text = monades.Items(0).ToString

        ComboFifo.Text = ComboFifo.Items(0).ToString
        GDB.Open(gConnect)
        '  GDB.BeginTrans()

        FILLComboBox("select EPO+';'+AFM,space(30)+AFM FROM PEL WHERE EIDOS='e' order by EPO ", pel)

    End Sub

    Private Sub POSO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles POSO.KeyUp
        FillListBox("SELECT STR(PARTIDA)+'; '+CONVERT(CHAR(10),HME,3)+';'+ STR(YPOL)+SPACE(50)+';'+ STR(ID),ID   FROM PARTIDES WHERE YPOL>0 AND KOD='" + EIDKOD.Text + "' ORDER BY HME", PARTIDES)
    End Sub

    Private Sub POSO_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POSO.TextChanged

    End Sub

    Private Sub PARTIDES_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PARTIDES.SelectedIndexChanged
        Dim s As Long = 0
        Dim k As Integer

        For k = 0 To PARTIDES.SelectedItems.Count - 1
            s = s + Val(Split(PARTIDES.SelectedItems(k).ToString(), ";")(2))
            ' Split(ListBox1.Items(N).ToString, ";")(0)
        Next
        synoltem.Text = s

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\symorfosi.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sh1")
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value




        Dim YLIKA As New DataTable

        ExecuteSQLQuery("SELECT PARTIDA,HME,ATIM,POSO  FROM TIMSPOL where HME='" + Format(HME.Value, "MM/dd/yyyy") + "' AND ATIM='" + ATIM.Text + "' ", YLIKA)

        xlWorkSheet.Cells(8, 3) = pel.Text
        xlWorkSheet.Cells(32, 2) = pel.Text
        xlWorkSheet.Cells(9, 8) = "--" + Format(Now, "dd/MM/yyyy")
        For KY = 0 To YLIKA.Rows.Count - 1
            'YL = YLIKA.Rows(KY).Item(0).ToString  'ΚΩΔΙΚΟΣ ΣΥΣΤΑΤΙΚΟΥ
            xlWorkSheet.Cells(15 + KY, 2) = KY + 1

            xlWorkSheet.Cells(15 + KY, 3) = YLIKA.Rows(KY).Item("PARTIDA").ToString
            xlWorkSheet.Cells(15 + KY, 4) = YLIKA.Rows(KY).Item("HME").ToString
            xlWorkSheet.Cells(15 + KY, 5) = YLIKA.Rows(KY).Item("POSO").ToString

            xlWorkSheet.Cells(15 + KY, 6) = "ΤΙΜ/ΔΑ"
            xlWorkSheet.Cells(15 + KY, 7) = YLIKA.Rows(KY).Item("atim").ToString


        Next

        For KY = YLIKA.Rows.Count To 10
            xlWorkSheet.Cells(15 + KY, 1) = ""
            xlWorkSheet.Cells(15 + KY, 2) = ""
            xlWorkSheet.Cells(15 + KY, 3) = ""
            xlWorkSheet.Cells(15 + KY, 4) = ""
            xlWorkSheet.Cells(15 + KY, 5) = ""
            xlWorkSheet.Cells(15 + KY, 6) = ""
            xlWorkSheet.Cells(15 + KY, 7) = ""

        Next


        xlWorkBook.Save()

        xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)








    End Sub


    Private Sub diagrafh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles diagrafh.Click
        Dim idpart As String
        Dim tem As String
        Dim IDtIMOLS As String
        For K = 0 To LINES.SelectedItems.Count - 1
            idpart = Split(LINES.SelectedItems(K).ToString(), ";")(4)
            IDtIMOLS = Split(LINES.SelectedItems(K).ToString(), ";")(5)
            tem = Split(LINES.SelectedItems(K).ToString(), ";")(2)
            GDB.Execute("update PARTIDES SET YPOL=YPOL+" + tem + " WHERE ID=" + idpart)
            GDB.Execute("DELETE FROM TIMSPOL WHERE ID=" + IDtIMOLS)
            'LINES.SelectedItems.Remove(LINES.SelectedItems.Item(i))
            LINES.Items.RemoveAt(LINES.SelectedItems(K))
        Next
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ANS = MsgBox("Είναι η Γραμογράφηση Τιμ.πώλησης ΟΚ  ??", MsgBoxStyle.YesNo)
        If ANS = MsgBoxResult.No Then
            Exit Sub
        End If


        OpenFileDialog1.ShowDialog()
        Dim ok As Boolean = True
        import_timologia(True, ok)
        If ok Then
            import_timologia(False, ok)
        Else
            MsgBox("δεν αποθηκεύθηκε το Excel")
        End If


    End Sub
    Private Sub import_timologia(ByVal check_only As Boolean, ByRef is_ok As Boolean)

        Dim N As Integer = 11



        ' Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim cPel As String
        Dim cEID As String
        Dim mHME As String
        Dim mAtim As String


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


        Dim D As String
        Dim flagPel As Integer
        Dim flagEID As Integer
        Dim mposo As Double
        Dim sql As String
        Dim merror As String
        Dim mon As String

        Dim mFiFo As String
        If Mid(ComboFifo.Text, 1, 1) = 1 Or Mid(ComboFifo.Text, 1, 1) = 3 Then
            mFiFo = ""
        Else
            mFiFo = " desc"
        End If

        Dim mScan As Integer
        mScan = 1

        ' If Val(Mid(ComboFifo.Text, 1, 1)) > 2 Then mScan = 1 Else mScan = 0
        '1.FIFO (Ανάλωση από Παλαιά Παρτίδα )
        '2.LIFO (Ανάλωση από Τελευταία Παρτίδα )
        '3:      .FIFO(+scanner)
        '4:      .LIFO(+scanner)



        ' GDB.BeginTrans()
        Dim w As Integer
        Dim NoPistotiko As Integer = 1

        Dim parastatiko As String = xlWorkSheet.Cells(9, 2).VALUE.ToString

        ' σβησιμο αλλαγης σελίδας   xlWorkSheet.range("b20").VALUE =   xlWorkSheet.Cells(20, 2).VALUE
        Dim nn As Integer
        'Dim cnn As String = ""
        For nn = xlWorkSheet.UsedRange.Rows.Count To 11 Step -1



            If xlWorkSheet.Cells(nn, 18).Value = Nothing Then

            Else

                If xlWorkSheet.Cells(nn, 18).Value.ToString = "Σελίδα" Then
                    xlWorkSheet.Rows(nn - 4 & ":" & nn).Delete()

                End If


            End If

        Next
        xlWorkBook.Save()



        'Exit Sub













        Do

            'parastatiko = xlWorkSheet.Cells(9, 2).VALUE.ToString
            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString

            If N > xlWorkSheet.UsedRange.Rows.Count Then
                If xlWorkSheet.Cells(N, 2).VALUE = Nothing Then
                    Exit Do
                End If
            End If


            If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 4).VALUE) Then
                If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 2).VALUE) = False Then
                    parastatiko = xlWorkSheet.Cells(N, 2).VALUE.ToString
                End If
            End If




            Try


                'ΨΑΧΝΩ ΝΑ ΔΡΩ ΤΟΝ ΚΩΔΙΚΟ ΤΟΥ pelath
                flagPel = 0
                w = 1

                While True  ' Not xlWorkSheet.Cells(N, 2).VALUE = Nothing

                    If N > xlWorkSheet.UsedRange.Rows.Count Then
                        If xlWorkSheet.Cells(N, 4).VALUE = Nothing Then
                            Exit While
                        End If
                    End If


                    If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 4).VALUE) Then
                        If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 2).VALUE) = False Then
                            If Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 1, 4) = "ΤΕΧΝ" Then
                            Else
                                parastatiko = xlWorkSheet.Cells(N, 2).VALUE.ToString
                            End If

                        End If
                    End If









                    w = 2

                    If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 4).VALUE) Then
                        cPel = "ΩΩΩΩ"
                    Else
                        cPel = xlWorkSheet.Cells(N, 4).VALUE.ToString

                    End If
                    w = 3
                    'βρηκα πελατη
                    If IsNumeric(Mid(cPel, 1, 6)) Then

                        mHME = Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10)
                        mHME = Mid(mHME, 4, 2) + "/" + Mid(mHME, 1, 2) + "/" + Mid(mHME, 7, 4)
                        mAtim = xlWorkSheet.Cells(N, 3).VALUE.ToString
                        If InStr(mAtim, " ") > 0 Then
                            mAtim = Trim(Mid(mAtim, InStr(mAtim, " "), 10))
                        End If

                        For N45 As Integer = 1 To 6
                            If Len(mAtim) > 0 Then
                                If Mid(mAtim, 1, 1) = "0" Then
                                    mAtim = Mid(mAtim, 2, Len(mAtim) - 1)
                                End If
                            End If

                        Next




                        Dim MFI As New DataTable
                        ExecuteSQLQuery("select count(*) from TIMSPOL WHERE ATIM='" + mAtim + "' AND HME='" + mHME + "'", MFI)

                        If MFI(0)(0) > 0 Then
                            MsgBox("ΥΠΑΡΧΕΙ ΗΔΗ ΤΟ ΤΙΜΟΛΟΓΙΟ " + mAtim)
                            is_ok = False
                            Exit Sub

                        End If













                        If InStr(mAtim, "ΠΙΣ") > 0 Or InStr(parastatiko, "ΔΠΑ") > 0 Or InStr(parastatiko, "ΔΠΕ") > 0 Then
                            NoPistotiko = -1
                        Else
                            NoPistotiko = 1
                        End If

                        w = 4
                        flagPel = 1
                        'line=
                        N = N + 1
                        Exit While
                    End If
                    N = N + 1
                End While


                'ΔΕΝ ΒΡΗΚΑ ΠΕΛΑΤΗ ΒΓΕΣ ΑΠΟ LOOP
                If flagPel = 0 Then
                    Exit Do
                End If


                'ΨΑΧΝΩ ΝΑ BΡΩ ΤΟYΣ ΚΩΔΙΚΟΥΣ ΤΩΝ  ΕΙΔΩΝ
                flagEID = 0




                ' LOOP IDIOY  TIMOLOGIOY
                While True '  Not xlWorkSheet.Cells(N, 2).VALUE = Nothing

                    w = 5
                    'If N > xlWorkSheet.UsedRange.Rows.Count Then
                    If xlWorkSheet.Cells(N, 2).VALUE = Nothing Then
                        Exit While
                    End If
                    cEID = Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 8, 6)
                    'δεβθγ νεχτ 4 λινεσ
                    If cEID = "000513" Then
                        w = 5
                    End If






                    w = 6
                    If Not IsNumeric(cEID) Then  '==============================================IsNumeric(cEID) ==============================================
                        Exit While
                    End If
                    Try

                        'DEBUG GDB.BeginTrans()

                        Dim isA_YLH As Integer = 0
                        Dim MaxIdTimsAnal As Long
                        Dim FFD2 As New DataTable
                        ExecuteSQLQuery("select max(ID) FROM TIMSANAL", FFD2)
                        If IsDBNull(FFD2(0)(0)) Then
                            MaxIdTimsAnal = 0
                        Else
                            MaxIdTimsAnal = FFD2(0)(0)
                        End If







                        ' cEID = xlWorkSheet.Cells(N, 2).VALUE.ToString
                        'If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 14).VALUE) Then
                        'mposo = Val(Kau_Aritmoy(xlWorkSheet.Cells(N, 9).VALUE.ToString))
                        'Else
                        If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 9).VALUE) Then
                            mposo = NoPistotiko * Val(Kau_Aritmoy(xlWorkSheet.Cells(N, 10).VALUE.ToString))
                        Else
                            mposo = NoPistotiko * Val(Kau_Aritmoy(xlWorkSheet.Cells(N, 9).VALUE.ToString))
                        End If

                        'End If
                        w = 7
                        Dim r22 As New ADODB.Recordset

                        r22.Open("select * FROM YLIKA WHERE KOD='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        w = 8
                        If r22.EOF Then
                            MsgBox(" ΔΕΝ ΥΠΑΡΧΕΙ Ο ΚΩΔΙΚΟΣ " + cEID + ". ΔΙΑΚΟΠΗΚΕ Η ΔΙΑΔΙΚΑΣΙΑ")
                            GDB.Execute("INSERT INTO YLIKA (KOD,N1) VALUES ('" + cEID + "',4)")
                            '   merror = 1
                            '  Exit Do

                        Else
                            'GDB.Execute("INSERT INTO TIMS (C1,KOD,PROM,ATIM,HME,POSO,YPOL) VALUES ('" + Mid(cPel, 1, 6) + "','" + cEID + "','" + Mid(cPel, 8, 16) + "','" + mAtim + "','" + mHME + "'," + mposo + "," + mposo + ")")
                        End If
                        w = 9
                        mon = r22("c1").Value.ToString
                        ' If r.EOF Then

                        'End If
                        r22.Close()
                        w = 10
                        If Val(Mid(mon, 1, 3)) > 0 Then
                            mposo = Replace(Str(Val(mposo) * 100), ",", ".")
                        End If


                        w = 11


                        'ΕΔΩ ΘΑ ΒΑΛΩ ΤΙΣ ΠΑΡΤΙΔΕΣ ΤΟΥ ΣΚΑΝΕΡ
                        'ΒΡΙΣΚΩ ΤΙΣ ΠΑΡΤΙΔΕΣ ΠΟΥ ΣΗΜΑΔΕΨΑ ΓΙΑ ΤΟ ΤΙΜΟΛΟΓΙΟ

                        Dim SQLPALET = "DROP TABLE DOKPALETTIM; "


                        If check_only = False Then
                            If mScan = 1 Then
                                Try
                                    GDB.Execute(SQLPALET)
                                Catch ex As Exception

                                End Try

                            End If
                        End If




                        SQLPALET = "SELECT   [ATIM],PARTIDA,SUM(P.POSO) AS SPOSO INTO DOKPALETTIM FROM [TECHNOPLASTIKI].[dbo].[PALETTIM] T "
                        SQLPALET = SQLPALET + " INNER JOIN PALETES P ON  SUBSTRING(SCAN,13,7)=P.PALET "
                        SQLPALET = SQLPALET + "WHERE  CONVERT(INTEGER,ATIM)=CONVERT(INT,'" + mAtim + "')   GROUP BY ATIM,PARTIDA;"
                        'ATIM	PARTIDA  	SPOSO
                        '012645	 1708206	2640
                        SQLPALET = SQLPALET + "UPDATE PARTIDES SET CH2=CONVERT(CHAR(10),(SELECT TOP 1 SPOSO FROM DOKPALETTIM WHERE PARTIDA=PARTIDES.PARTIDA ) )   WHERE  PARTIDA IN (SELECT PARTIDA FROM DOKPALETTIM)"




                        If check_only = False Then
                            If mScan = 1 Then
                                GDB.Execute(SQLPALET)
                            End If
                        End If



                        Dim PosoReal As Long
                        'ελεγχος παρτίδων
                        '"SELECT STR(PARTIDA)+'; '+CONVERT(CHAR(10),HME,3)+';'+ STR(YPOL)+SPACE(50)+';'+ STR(ID),ID   FROM PARTIDES WHERE YPOL>0 AND KOD='" + EIDKOD.Text + "' ORDER BY HME
                        'r.Open("select * from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' ORDER BY HME " + mFiFo, GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                        Dim mSQLscan As String
                        If mScan = 1 Then mSQLscan = " CH2 DESC, " Else mSQLscan = " "



                        'Αυτά που έχω σημαδέψει με το πιστολάκι δεν φτάνουν το ποσό του τιμολογίου (mposo) .
                        ' Υπάρχουν αρκετά ασημάδευτα (ch2 is null) για να συμπληρώσουν ?
                        ' Αν ναι έχει καλώς
                        'Αν όχι σημαίνει ότι στις ίδιες σειρές που εχω μαρκαρει (ch2<>null) πρέπει να πάρω όλη την ποσότητα
                        ' γιατί δεν θα συμπληρώσω το τιμολόγιο
                        ' αγνοώ δηλαδή το πιστόλι και πάω lifo ή fifo
                        'παράδειγμα που έτυχε

                        'PARTIDA	HME	KOD	TIMOLOGIA	TEMAXIA	ID	N1	N2	CH1	CH2	                                                    ch2	        ypol
                        '1708546	2021-05-13 00:00:00.000	001299	001319;1485-2020;2020-12-21*	7850	36802	7762	0	NULL	 7800      	7850	    
                        '1708547	2021-05-13 00:00:00.000	001299	001319;1597-2020;2021-04-06*	7750	36803	7998	230	NULL	 5850      	7750	      
                        '1708550	2021-05-14 00:00:00.000	001299	001319;1597-2020;2021-04-06*	17550	36806	7998	233	NULL	15600      17550
                        '                                                                                                               -----       -----
                        '                                                                                                               29250      33150       
                        ' αν ακολουθησω το πιστολάκι μετά δεν θα εχει αλλη παρτιδα να παρω και θα μείνει λειψό το τιμολογιο
                        Dim rr1 As New DataTable, rr2 As New DataTable, SMARKAR As Single, SUNMARK As Single
                        ExecuteSQLQuery("select  ISNULL( sum(YPOL),0) AS YP,ISNULL(SUM(CONVERT(REAL,CH2)),0)  AS MCH2 from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' and (NOT CH2 IS NULL)", rr1)
                        SMARKAR = rr1(0)(1)
                        ExecuteSQLQuery("select  ISNULL( sum(YPOL),0) AS YP from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' and ( CH2 IS NULL)", rr2)
                        SUNMARK = rr2(0)(0)

                        ' ΤΑ ΑΜΑΡΚΑΡΙΣΤΑ ΔΕΝ ΦΤΑΝΟΥΝ ΝΑ ΚΑΛΥΨΟΥΝ ΑΥΤΑ ΠΟΥ ΜΑΣ ΛΕΙΠΟΥΝ ΤΟΤΕ ΠΑΩ LIFO/FIFO ΚΑΙ ΑΓΝΟΩ ΤΟ ΜΑΡΚΑΡΙΑΣΜΑ

                        Dim rr As New DataTable
                        If SUNMARK < (mposo - SMARKAR) Then
                            ExecuteSQLQuery("select *,'' AS MCH2 from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' ORDER BY " + mSQLscan + "HME " + mFiFo, rr)


                        Else
                            ExecuteSQLQuery("select *,isnull(CH2,'') AS MCH2 from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' ORDER BY " + mSQLscan + "HME " + mFiFo, rr)

                        End If




                    
                        w = 111
                        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                        If NoPistotiko = -1 Then  'ΠΙΣΤΩΤΙΚΑ ΕΠΙΣΤΡΟΦΕΣ
                            ' PROSOXH ------------------------------
                            'ΣΤΑ ΠΙΣΤΩΤΙΚΑ ΣΤΟ TIMSPOL ΑΠΟΘΗΚΕΥΩ ΤΑ ΕΞΗΣ ΓΙΑ ΝΑ ΜΠΟΡΩ ΝΑ ΤΑ ΣΒΗΣΩ
                            'C1='ΠΙΣ'
                            'Ν1= ID PARTIDES
                            'N2 = ID TIMS 




                            Dim RSYN2 As New ADODB.Recordset
                            RSYN2.Open("SELECT COUNT(*) FROM SYNTAGES WHERE KOD='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            'cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
                            Dim mIDpart As Long, mIDtimagor As Long

                            If RSYN2(0).Value > 0 Then  ' EINAI ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ
                                If check_only = False Then
                                    Dim mp As Long, partides As New DataTable
                                    ExecuteSQLQuery("SELECT max(PARTIDA) FROM PARTIDES  WHERE PARTIDA LIKE '17%' ", partides)
                                    If IsDBNull(partides(0)(0)) Then
                                        mp = 1700000
                                    Else
                                        mp = partides(0)(0) + 1
                                    End If
                                    Dim HM As String = Format(Now, "MM/dd/yyyy") ' 
                                    Dim SQLQuery As String = "INSERT INTO PARTIDES (PARTIDA,HME,KOD,TIMOLOGIA,TEMAXIA,YPOL,N1) VALUES(" + Str(mp) + ",'" + HM + "','" + cEID + "','" + mAtim + "'," + Replace(Str(Math.Abs(mposo)), ",", ".") + "," + Replace(Str(Math.Abs(mposo)), ",", ".") + ",0)"
                                    GDB.Execute(SQLQuery)
                                    ExecuteSQLQuery("SELECT max(ID) FROM PARTIDES  WHERE PARTIDA LIKE '17%' ", partides)
                                    If IsDBNull(partides(0)(0)) Then
                                    Else
                                        mIDpart = partides(0)(0)
                                        mIDtimagor = 0
                                    End If

                                    'GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(POS2.Rows(LL)("YPOL")), ",", ".") + ",'" + cEID + "','" + mAtim + "')")
                                    'GDB.Execute("UPDATE TIMS SET YPOL=0 WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                                End If
                            Else  ' YLIKO 
                                If check_only = False Then


                                    sql = "insert into TIMS (HME,POSO,ATIM,KOD,PROM,YPOL) VALUES("
                                    sql = sql + "'" + mHME + "',"
                                    sql = sql + Replace(Str(Math.Abs(mposo)), ",", ".") + ","
                                    sql = sql + "'" + mAtim + "',"
                                    sql = sql + "'" + cEID + "',"
                                    sql = sql + "'" + cPel + "',"
                                    sql = sql + Replace(Str(Math.Abs(mposo)), ",", ".") + ")"
                                    ExecuteSQLQuery(sql)
                                    Dim PARTIDES As New DataTable
                                    ExecuteSQLQuery("SELECT max(ID) FROM TIMS ", PARTIDES)
                                    If IsDBNull(PARTIDES(0)(0)) Then
                                    Else
                                        mIDpart = 0
                                        mIDtimagor = PARTIDES(0)(0)
                                    End If








                                    'GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(POS2.Rows(LL)("YPOL")), ",", ".") + ",'" + cEID + "','" + mAtim + "')")
                                    'GDB.Execute("UPDATE TIMS SET YPOL=0 WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                                End If
                            End If
                            RSYN2.Close()
                            If check_only = False Then
                                sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1,N2,C1) VALUES("
                                sql = sql + "'" + mHME + "',"
                                sql = sql + Replace(Str(Math.Abs(mposo)), ",", ".") + ","
                                sql = sql + "'" + mAtim + "',"
                                sql = sql + "'" + cEID + "',"
                                sql = sql + "'" + Mid(cPel, 1, 6) + "',"
                                sql = sql + "0" + ",'" + "." + "'," + Str(mIDpart) + "," + Str(mIDtimagor) + ",'ΠΙΣ' )"
                                w = 114
                                ' r.Close()

                                'If check_only = False Then
                                GDB.Execute(sql)
                            End If


                            'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            ' χρεωστικα παραστατικά
                        Else


                            'ΔΕΝ ΒΡΗΚΑ ΚΑΘΟΛΟΥ ΠΑΡΤΙΔΕΣ ΣΕ ΑΥΤΟ ΤΟ ΕΙΔΟΣ
                            'NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN
                            'If r.EOF Then
                            If rr.Rows.Count = 0 Then
                                'ΕΙΝΑΙ ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ????
                                Dim RSYN As New ADODB.Recordset
                                RSYN.Open("SELECT COUNT(*) FROM SYNTAGES WHERE KOD='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)




                                'cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
                                If RSYN(0).Value > 0 Then  ' EINAI ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ
                                    LINES.Items.Add(";;;;ΔΕΝ ΥΠΑΡΧΕΙ ΠΑΡΤΙΔΑ ΓΙΑ ΤΟ ΕΙΔΟΣ " + cEID)
                                    MsgBox(" δεν υπαρχει καθόλου υπολοιπο παρτιδων για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                    is_ok = False
                                    Exit While
                                    'Exit Sub

                                    w = 112
                                    'sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
                                    'sql = sql + "'" + mHME + "',"
                                    'sql = sql + Kau_Aritmoy(mposo) + ","
                                    'sql = sql + "'" + mAtim + "',"
                                    'sql = sql + "'" + cEID + "',"
                                    'sql = sql + "'" + Mid(cPel, 1, 6) + "',"
                                    'sql = sql + "0" + ",'" + "." + "'," + "0" + " )"
                                    'w = 114
                                    '' r.Close()
                                    'Try
                                    '    GDB.Execute(sql)
                                    '    LINES.Items.Add("Προιόν;" + cEID + ";Τιμολ.;" + mAtim + ";" + Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10) + ";Παρτίδα;0;" + cPel)
                                    'Catch ex As Exception
                                    '    MsgBox("δεν αποθηκευτηκε το τιμολογιο " + mAtim + "  " + ex.Message)
                                    'End Try

                                    'cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
                                Else  'ΔΕΝ ΕΙΝΑΙ  ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ  ΑΡΑ ΕΙΝΑΙ ΠΡΩΤΗ ΥΛΗ Ή ΒΟΗΘΗΤΙΚΗ
                                    ' ΒΡΕΣ ΤΑ ΤΙΜΟΛΟΓΙΑ ΠΟΥ ΕΧΟΥΝ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ

                                    isA_YLH = 1
                                    Dim POS2 As New DataTable
                                    Dim r2 As New ADODB.Recordset
                                    r2.Open("select sum(YPOL) FROM TIMS WHERE RTRIM(LTRIM(KOD))='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                    If IsDBNull(r2(0).Value) Then
                                        ' < mposo Then
                                        MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο TIM.AΓOΡAΣ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                        is_ok = False
                                        Exit Sub
                                    Else
                                        If r2(0).Value < mposo Then
                                            MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο TIM.AΓOΡAΣ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                            is_ok = False
                                            Exit Sub
                                        End If



                                    End If
                                    r2.Close()



                                    ExecuteSQLQuery("SELECT  KOD,YPOL,ATIM,HME,ID  FROM TIMS where YPOL>0 AND RTRIM(LTRIM(KOD))='" + cEID + "' ORDER BY HME " + mFiFo, POS2)
                                    If POS2.Rows.Count = 0 Then  ' '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Else '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim LL As Integer, m_MPOSO As Single = mposo
                                        '=======///==============
                                        For LL = 0 To POS2.Rows.Count - 1
                                            'ΤΟ ΠΟΣΟ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΕΙ ΕΙΝΑΙ ΛΙΓΟΤΕΡΟ ΑΠΟ ΤΟ ΤΙΜΟΛ.ΑΓΟΡΑΣ
                                            'ΟΠΟΤΕ ΑΦΑΙΡΩ ΟΛΟ ΤΟ ΠΟΣΟ ΤΟΥ ΤΙΜΟΛ.ΠΩΛΗΣΗΣ (M_POSO)
                                            If m_MPOSO < POS2.Rows(LL)("YPOL") Then
                                                'DEBUG
                                                If check_only = False Then
                                                    GDB.Execute("UPDATE TIMS SET YPOL=YPOL-" + Replace(Str(m_MPOSO), ",", ".") + " WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                                                    GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(m_MPOSO), ",", ".") + ",'" + cEID + "','" + mAtim + "')")
                                                End If

                                                m_MPOSO = 0

                                            Else '  m_MPOSO > POS2.Rows(LL)("YPOL") Then
                                                'ΤΟ ΠΟΣΟ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΕΙ ΕΙΝΑΙ MEGALYTERO ΑΠΟ ΤΟ ΤΙΜΟΛ.ΑΓΟΡΑΣ
                                                'ΟΠΟΤΕ ΑΦΑΙΡΩ ΟΛΟ ΤΟ ΠΟΣΟ ΤΟΥ ΤΙΜΟΛ.AGORAS POS2.Rows(LL)("YPOL")
                                                m_MPOSO = m_MPOSO - POS2.Rows(LL)("YPOL")
                                                'DEBUG
                                                If check_only = False Then
                                                    GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(POS2.Rows(LL)("YPOL")), ",", ".") + ",'" + cEID + "','" + mAtim + "')")
                                                    GDB.Execute("UPDATE TIMS SET YPOL=0 WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                                                End If

                                            End If
                                            If m_MPOSO = 0 Then
                                                Exit For
                                            End If
                                        Next
                                        '========///=============
                                    End If  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    '========================
                                    sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
                                    sql = sql + "'" + mHME + "',"
                                    sql = sql + Kau_Aritmoy(mposo) + ","
                                    sql = sql + "'" + mAtim + "',"
                                    sql = sql + "'" + cEID + "',"
                                    sql = sql + "'" + Mid(cPel, 1, 6) + "',"
                                    sql = sql + "0" + ",'" + "." + "'," + "0" + " )"
                                    w = 114
                                    ' r.Close()
                                    Try
                                        If check_only = False Then
                                            GDB.Execute(sql)
                                        End If

                                        LINES.Items.Add("Α ΥΛΗ " + cEID + ";Τιμολ.;" + mAtim + ";" + Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10) + ";ANAΛΩΣΗ ΑΠΟ ΤΙΜΟΛΟΓΙΑ ΑΓΟΡΑΣ;" + cPel)
                                    Catch ex As Exception
                                        MsgBox("δεν αποθηκευτηκε το τιμολογιο " + mAtim + "  " + ex.Message)
                                    End Try
                                End If
                                'cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
                                w = 115

                                'End If  'NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN
                            Else ' ΒΡΗΚΑ ΠΑΡΤΙΔΕΣ

                                Dim POS2 As New DataTable
                                Dim r2 As New ADODB.Recordset
                                r2.Open("select sum(YPOL) FROM  PARTIDES WHERE RTRIM(LTRIM(KOD))='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                                If IsDBNull(r2(0).Value) Then
                                    ' < mposo Then
                                    MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο ΠΑΡΤΙΔΩΝ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                    is_ok = False
                                    Exit While
                                    'Exit Sub
                                Else
                                    If r2(0).Value < mposo Then
                                        MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο ΠΑΡΤΙΔΩΝ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                        is_ok = False
                                        Exit While
                                        'Exit Sub
                                    End If
                                End If

                                'If r2(0).Value < mposo Then
                                '    MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο ΠΑΡΤΙΔΩΝ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                '    Exit Sub

                                'End If
                                r2.Close()



                                'Do While Not R.EOF
                                For LL = 0 To (rr.Rows.Count - 1)
                                    w = 12
                                    'ΤΡΑΒΑΩ ΑΠΟ ΤΗΝ ΠΡΩΤΗ ΠΑΛΙΟΤΕΡΗ ΠΑΡΤΙΔΑ
                                    If rr(LL)("YPOL") > (mposo) Then
                                        'PosoReal = (mposo)
                                        If Val(rr(LL)("MCH2")) > 0 Then
                                            PosoReal = Val(rr(LL)("mch2"))
                                        Else
                                            PosoReal = (mposo)
                                        End If




                                    Else
                                        If Val(rr(LL)("mch2")) > 0 Then
                                            PosoReal = Val(rr(LL)("mch2"))
                                        Else
                                            PosoReal = rr(LL)("YPOL")
                                        End If


                                    End If
                                    w = 13
                                    sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
                                    sql = sql + "'" + mHME + "',"
                                    sql = sql + Kau_Aritmoy(Str(PosoReal)) + ","
                                    sql = sql + "'" + mAtim + "',"
                                    sql = sql + "'" + cEID + "',"
                                    sql = sql + "'" + Mid(cPel, 1, 6) + "',"
                                    sql = sql + rr(LL)("partida").ToString + ",'" + "." + "'," + "0" + " )"
                                    w = 14
                                    If check_only = False Then
                                        GDB.Execute(sql)
                                    End If

                                    w = 15
                                    'ExecuteSQLQuery("select max(ID) FROM TIMSPOL")
                                    ' idtimols = sqlDT.Rows(0)(0).ToString



                                    'DEBUG
                                    If check_only = False Then
                                        GDB.Execute("UPDATE PARTIDES SET YPOL=YPOL-" + Kau_Aritmoy(Str(PosoReal)) + " WHERE PARTIDA=" + rr(LL)("partida").ToString)
                                        If mScan = 1 Then
                                            'TEST WORKING  UPDATE  [TECHNOPLASTIKI].[dbo].[PARTIDES] SET CH2=CONVERT(CHAR(10), CONVERT(REAL,CH2)+200)    WHERE PARTIDA='1708206'
                                            GDB.Execute("UPDATE PARTIDES SET CH2=CONVERT(CHAR(10), CONVERT(REAL,CH2) - " + Kau_Aritmoy(Str(PosoReal)) + ")  WHERE PARTIDA=" + rr(LL)("partida").ToString)

                                            'Dim FFD As New DataTable, mLastID As Integer
                                            'ExecuteSQLQuery("select max(ID) FROM TIMSPOL", FFD)
                                            ' mLastID = FFD(0)(0)
                                            ' ExecuteSQLQuery("update PALETES SET select max(ID) FROM TIMSPOL", FFD)






                                        End If
                                    End If




                                    w = 16
                                    LINES.Items.Add("Προιόν;" + cEID + ";Τιμολ.;" + Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10) + ";" + mAtim + ";Παρτίδα;" + rr(LL)("partida").ToString + ";" + cPel)
                                    'Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10)




                                    If PosoReal >= Val(mposo) Then
                                        Exit For
                                    Else
                                        'αφαιρω αυτό που παρτιδοποιήθηκε
                                        mposo = mposo - PosoReal

                                    End If
                                    ' r.MoveNext()
                                    w = 17

                                Next
                                ' Loop


                                ' r.Close()

                            End If  ' ΒΡΗΚΑ ΠΑΡΤΙΔΕΣ
                            'End If  'NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN

                        End If 'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM    NoPistotiko = -1

                        w = 18
                        flagPel = 1

                        'DEBUG GDB.CommitTrans()

                        'ΕΝΗΜΕΡΩΝΩ ΤΑ TIMSANAL ME TO ID TOY TIMSPOL ΠΟΥ ΔΗΜΙΟΥΡΓΗΘΗΚΕ
                        If isA_YLH = 1 Then
                            Dim maxIDTimspol As Long
                            Dim FFD As New DataTable
                            ExecuteSQLQuery("select max(ID) FROM TIMSPOL", FFD)
                            maxIDTimspol = FFD(0)(0)
                            If check_only = False Then
                                ExecuteSQLQuery("update TIMSANAL set N2=" + Str(maxIDTimspol) + " WHERE N2=-3000 AND ID>" + Str(MaxIdTimsAnal))
                            End If

                        End If




                    Catch ex As Exception
                        MsgBox("λαθος στην σειρά " + Str(N) + Chr(13) + Err.Description)
                        xlWorkBook.Close()
                        xlApp.Quit()

                        releaseObject(xlApp)
                        releaseObject(xlWorkBook)
                        releaseObject(xlWorkSheet)

                        If check_only = False Then
                            MsgBox("ΑΠΟΘΗΚΕΥUHKAN TA TIMOLOGIA. ΜΕΧΡΙ ΣΕΙΡΑ EXCEL " + Str(N))
                        End If
                        'DEBUG GDB.RollbackTrans()

                        Exit Sub

                    End Try



                    'Else ' ΤΕΛΕΙΩΣΕ ΤΟ ΤΙΜΟΛΟΓΙΟ ΚΑΙ ΠΑΩ ΠΑΡΑΚΑΤΩ  '============================================================================================
                    '    Exit While

                    'End If '============================================================================================
                    N = N + 1
                End While   ' λοοπ mesa sto idio timologio


                If check_only = False Then
                    If mScan = 1 Then
                        GDB.Execute("UPDATE [PARTIDES] SET CH2=NULL ")
                    End If

                End If



                ' GDB.Execute("update SYNTAGES SET POSOSTO=" + Replace(Str(xlWorkSheet.Cells(N, 11).VALUE / 1), ",", ".") + " WHERE  KOD='" + cPel + "' AND KODSYNOD='" + D + "'")
            Catch ex As Exception
                MsgBox("κωδ.λαθους w=" + Str(w) + " ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + Str(N) + " " + cPel + " == " + D + "--" + ex.Message)
                'Exit Do

                MsgBox("λαθος στην σειρά " + Str(N) + Chr(13) + Err.Description)
                xlWorkBook.Close()
                xlApp.Quit()

                releaseObject(xlApp)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkSheet)

                If check_only = False Then
                    MsgBox("ΑΠΟΘΗΚΕΥUHKAN TA TIMOLOGIA. ΜΕΧΡΙ ΣΕΙΡΑ EXCEL " + Str(N))
                End If
                'DEBUG GDB.RollbackTrans()

                Exit Sub







            End Try


            N = N + 1
            Me.Text = N
        Loop Until False '  xlWorkSheet.Cells(N, 9).VALUE Is Nothing





        'ANS = MsgBox("ΝΑ ΑΠΟΘΗΚΕΥΤΟΥΝ ΟΙ ΕΙΣΑΓΩΓΕΣ ΤΙΜΟΛΟΓΙΩΝ (ΣΕΙΡΕΣ EXCEL " + Str(N) + " )  ", MsgBoxStyle.YesNo)

        'If ANS = MsgBoxResult.Yes Then
        '    GDB.CommitTrans()
        'Else
        '    GDB.RollbackTrans()
        'End If







        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        If check_only = False Then
            'ENHMERΩΝΕΙ ΤΙΣ ΠΑΛΕΤΕΣ ΜΕ ΤΟ ΤΙΜΟΛΟΓΙΟ ΠΟΥ ΠΗΓΕ Η ΚΑΘΕ ΠΑΛΕΤΑ
            If mScan = 1 Then
                GDB.Execute("UPDATE PALETES  SET ATIMPOL=(SELECT TOP 1 ATIM FROM PALETTIM WHERE  CONVERT(INT,SUBSTRING(SCAN,13,7))=PALET)  WHERE PALET IN (SELECT CONVERT(INT,SUBSTRING(SCAN,13,7)) FROM PALETTIM) ")
            End If

            MsgBox("ΑΠΟΘΗΚΕΥUHKAN TA TIMOLOGIA. ΜΕΧΡΙ ΣΕΙΡΑ EXCEL " + Str(N))
        End If



    End Sub

    Private Sub PYLONIMP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PYLONIMP.Click
        ' PYLONimport_timologia(True, 1)
        Dim fg As Integer
        fg = MsgBox("ΝΑ ΕΙΣΑΧΘΟΥΝ ΤΑ ΤΙΜΟΛΟΓΙΑ? ", MsgBoxStyle.YesNo)
        If fg = MsgBoxResult.No Then
            MsgBox("ΑΚΥΡΩΣΗ")
            Exit Sub
        End If

        Dim ok As Boolean = True
        PYLONimport_timologia(True, ok)
        If ok Then
            PYLONimport_timologia(False, ok)
        Else
            MsgBox("δεν αποθηκεύθηκε το Excel")
        End If





    End Sub

    Private Sub PYLONimport_timologia(ByVal check_only As Boolean, ByRef is_ok As Boolean)
        Dim mDAY, mMONTH, mYEAR As Integer
        MDAY = HME.Value.Day
        MMONTH = HME.Value.Month
        MYEAR = HME.Value.Year
        Dim N As Integer = 11

        Dim GPYL As New ADODB.Connection
        GPYL.Open("DSN=PYLONTECHNOPLASTIKI;uid=sa;pwd=p@ssw0rd")
        Dim R32 As New ADODB.Recordset

        Dim SQLAPOT As String = "select HEDOCCODE AS ATIM,HEDOCNUM,[A3].[HEITEMCODE] AS KODE  "
        SQLAPOT = SQLAPOT + " ,  [A3].[HEITEMDESCRIPTION] AS ONO ,  [A3].[HEAQTY] AS POSOTHTA, HEOFFICIALDATE as HME,"
        SQLAPOT = SQLAPOT + "'00000' AS KODIKOSPELATH,HEENTITYDESCR "
        SQLAPOT = SQLAPOT + " from [HEWENTLINES] [A3] WITH(NOLOCK)  inner join [HEWAREHOUSEENTRIES] [A2] WITH(NOLOCK)  on ([A3].[HEWENTID] = [A2].[HEID])"
        SQLAPOT = SQLAPOT + " inner join [HEDOCENTRIES] [A1] WITH(NOLOCK)  on ([A2].[HEDENTID] = [A1].[HEID] ) "
        SQLAPOT = SQLAPOT + " WHERE [A1].HEDOCCODE LIKE '%ΕΞΠ_%' AND DAY(HEOFFICIALDATE)=" + Str(mDAY) + " AND MONTH(HEOFFICIALDATE)=" + Str(mMONTH) + " AND YEAR(HEOFFICIALDATE)=" + Str(mYEAR) + " "


        Dim SQLP As String
        SQLP = " SELECT    HEDOCCODE AS ATIM,HEDOCNUM,[HEITEMCODE] AS KODE"
        SQLP = SQLP + ",[HEITEMDESCRIPTION] AS ONO ,[HEAQTY] AS [POSOTHTA],[HEOFFICIALDATE] AS HME,"
        SQLP = SQLP + "(SELECT HECODE FROM HECUSTOMERS WHERE HEID=C.HEBILLCSTMID) AS KODIKOSPELATH,C.HEENTITYDESCR  FROM [HECENTLINES] E "
        SQLP = SQLP + " INNER JOIN [HEDOCENTRIES] T  ON E.HEDENTID=T.HEID LEFT JOIN [HECOMMERCIALENTRIES] C ON E.HEDENTID=C.HEDENTID "
        SQLP = SQLP + "WHERE T.HESTATUS=0 AND LEFT(HEDOCCODE,4) IN ('ΔΑΒ-','ΕΞΠΟ','ΕΞΠΑ','ΤΔΑ-','ΔΑΠ-','ΔΑΧ-','ΔΑΥ-','ΔΕΠ-','ΔΕΠΧ','ΔΠΤ-') AND DAY(HEOFFICIALDATE)=" + Str(mDAY) + " AND MONTH(HEOFFICIALDATE)=" + Str(mMONTH) + " AND YEAR(HEOFFICIALDATE)=" + Str(mYEAR) + "  ORDER BY HME ,ATIM"


        R32.Open(SQLAPOT + " UNION " + SQLP, GPYL, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)






        ' Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim cPel As String
        Dim cEID As String
        Dim mHME As String
        Dim mAtim As String


        'Dim xlApp As Excel.Application
        'Dim xlWorkBook As Excel.Workbook
        'Dim xlWorkSheet As Excel.Worksheet

        'xlApp = New Excel.ApplicationClass
        'xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
        'xlWorkSheet = xlWorkBook.Worksheets(1)
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value


        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)


        Dim D As String
        Dim flagPel As Integer
        Dim flagEID As Integer
        Dim mposo As Double
        Dim sql As String
        Dim merror As String
        Dim mon As String

        Dim mFiFo As String
        If Mid(ComboFifo.Text, 1, 1) = 1 Or Mid(ComboFifo.Text, 1, 1) = 3 Then
            mFiFo = ""
        Else
            mFiFo = " desc"
        End If

        Dim mScan As Integer
        mScan = 1

        ' If Val(Mid(ComboFifo.Text, 1, 1)) > 2 Then mScan = 1 Else mScan = 0
        '1.FIFO (Ανάλωση από Παλαιά Παρτίδα )
        '2.LIFO (Ανάλωση από Τελευταία Παρτίδα )
        '3:      .FIFO(+scanner)
        '4:      .LIFO(+scanner)



        ' GDB.BeginTrans()
        Dim w As Integer
        Dim NoPistotiko As Integer = 1

        Dim parastatiko As String = R32("ATIM").Value.ToString   'xlWorkSheet.Cells(9, 2).VALUE.ToString

        ' σβησιμο αλλαγης σελίδας   xlWorkSheet.range("b20").VALUE =   xlWorkSheet.Cells(20, 2).VALUE
        Dim nn As Integer

        Do While Not R32.EOF

            'parastatiko = xlWorkSheet.Cells(9, 2).VALUE.ToString
            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString
            parastatiko = R32("ATIM").Value.ToString
            Try
                'ΨΑΧΝΩ ΝΑ ΔΡΩ ΤΟΝ ΚΩΔΙΚΟ ΤΟΥ pelath
                flagPel = 0
                w = 1
                '=======================================================================
                ' Not xlWorkSheet.Cells(N, 2).VALUE = Nothing
                cPel = R32("KODIKOSPELATH").Value.ToString
                mHME = Format(R32("HME").Value, "MM/dd/yyyy") '    .ToString
                mAtim = R32("ATIM").Value.ToString

                'βρηκα πελατη
                ' If IsNumeric(Mid(cPel, 1, 6)) Then
                Dim MFI As New DataTable
                ExecuteSQLQuery("select count(*) from TIMSPOL WHERE ATIM='" + mAtim + "' AND HME='" + mHME + "'", MFI)

                If MFI(0)(0) > 0 Then
                    MsgBox("ΥΠΑΡΧΕΙ ΗΔΗ ΤΟ ΤΙΜΟΛΟΓΙΟ " + mAtim)
                    is_ok = False
                    Exit Sub

                End If
                ''ΔΕΠ-','ΔΠΤ-'
                If InStr(mAtim, "ΔΕΠ") > 0 Or InStr(parastatiko, "ΔΠΤ") > 0 Then
                    NoPistotiko = -1
                Else
                    NoPistotiko = 1
                End If

                '=======================================================



                'ΨΑΧΝΩ ΝΑ BΡΩ ΤΟYΣ ΚΩΔΙΚΟΥΣ ΤΩΝ  ΕΙΔΩΝ
                flagEID = 0




                ' LOOP IDIOY  TIMOLOGIOY
                While R32("ATIM").Value.ToString = parastatiko '  Not xlWorkSheet.Cells(N, 2).VALUE = Nothing

                    w = 5

                    cEID = R32("KODE").Value.ToString ' Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 8, 6)
                    w = 6

                    Try

                        'DEBUG GDB.BeginTrans()

                        Dim isA_YLH As Integer = 0
                        Dim MaxIdTimsAnal As Long
                        Dim FFD2 As New DataTable
                        ExecuteSQLQuery("select max(ID) FROM TIMSANAL", FFD2)
                        If IsDBNull(FFD2(0)(0)) Then
                            MaxIdTimsAnal = 0
                        Else
                            MaxIdTimsAnal = FFD2(0)(0)
                        End If







                        ' cEID = xlWorkSheet.Cells(N, 2).VALUE.ToString
                        'If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 14).VALUE) Then
                        'mposo = Val(Kau_Aritmoy(xlWorkSheet.Cells(N, 9).VALUE.ToString))
                        'Else
                        'If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 9).VALUE) Then
                        '    mposo = NoPistotiko * Val(Kau_Aritmoy(xlWorkSheet.Cells(N, 10).VALUE.ToString))
                        'Else
                        '    mposo = NoPistotiko * Val(Kau_Aritmoy(xlWorkSheet.Cells(N, 9).VALUE.ToString))
                        'End If

                        mposo = NoPistotiko * R32("POSOTHTA").Value




                        'End If
                        w = 7
                        Dim r22 As New ADODB.Recordset

                        r22.Open("select * FROM YLIKA WHERE KOD='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        w = 8
                        If r22.EOF Then
                            MsgBox(" ΔΕΝ ΥΠΑΡΧΕΙ Ο ΚΩΔΙΚΟΣ " + cEID + ". ΔΙΑΚΟΠΗΚΕ Η ΔΙΑΔΙΚΑΣΙΑ")
                            GDB.Execute("INSERT INTO YLIKA (KOD,N1) VALUES ('" + cEID + "',4)")
                            '   merror = 1
                            '  Exit Do

                        Else
                            'GDB.Execute("INSERT INTO TIMS (C1,KOD,PROM,ATIM,HME,POSO,YPOL) VALUES ('" + Mid(cPel, 1, 6) + "','" + cEID + "','" + Mid(cPel, 8, 16) + "','" + mAtim + "','" + mHME + "'," + mposo + "," + mposo + ")")
                        End If
                        w = 9
                        mon = r22("c1").Value.ToString
                        ' If r.EOF Then

                        'End If
                        r22.Close()
                        w = 10
                        If Val(Mid(mon, 1, 3)) > 0 Then
                            mposo = Replace(Str(Val(mposo) * 100), ",", ".")
                        End If


                        w = 11


                        'ΕΔΩ ΘΑ ΒΑΛΩ ΤΙΣ ΠΑΡΤΙΔΕΣ ΤΟΥ ΣΚΑΝΕΡ
                        'ΒΡΙΣΚΩ ΤΙΣ ΠΑΡΤΙΔΕΣ ΠΟΥ ΣΗΜΑΔΕΨΑ ΓΙΑ ΤΟ ΤΙΜΟΛΟΓΙΟ

                        Dim SQLPALET = "DROP TABLE DOKPALETTIM; "


                        If check_only = False Then
                            If mScan = 1 Then
                                Try
                                    GDB.Execute(SQLPALET)
                                Catch ex As Exception

                                End Try

                            End If
                        End If




                        SQLPALET = "SELECT   [ATIM],PARTIDA,SUM(P.POSO) AS SPOSO INTO DOKPALETTIM FROM [TECHNOPLASTIKI].[dbo].[PALETTIM] T "
                        SQLPALET = SQLPALET + " INNER JOIN PALETES P ON  SUBSTRING(SCAN,13,7)=P.PALET "
                        SQLPALET = SQLPALET + "WHERE  CONVERT(INTEGER,ATIM)=CONVERT(INT,'" + Mid(mAtim, 5, 10) + "')   GROUP BY ATIM,PARTIDA;"
                        'ATIM	PARTIDA  	SPOSO
                        '012645	 1708206	2640
                        SQLPALET = SQLPALET + "UPDATE PARTIDES SET CH2=CONVERT(CHAR(10),(SELECT TOP 1 SPOSO FROM DOKPALETTIM WHERE PARTIDA=PARTIDES.PARTIDA ) )   WHERE  PARTIDA IN (SELECT PARTIDA FROM DOKPALETTIM)"




                        If check_only = False Then
                            If mScan = 1 Then
                                GDB.Execute(SQLPALET)
                            End If
                        End If



                        Dim PosoReal As Long
                        'ελεγχος παρτίδων
                        '"SELECT STR(PARTIDA)+'; '+CONVERT(CHAR(10),HME,3)+';'+ STR(YPOL)+SPACE(50)+';'+ STR(ID),ID   FROM PARTIDES WHERE YPOL>0 AND KOD='" + EIDKOD.Text + "' ORDER BY HME
                        'r.Open("select * from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' ORDER BY HME " + mFiFo, GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                        Dim mSQLscan As String
                        If mScan = 1 Then mSQLscan = " CH2 DESC, " Else mSQLscan = " "



                        'Αυτά που έχω σημαδέψει με το πιστολάκι δεν φτάνουν το ποσό του τιμολογίου (mposo) .
                        ' Υπάρχουν αρκετά ασημάδευτα (ch2 is null) για να συμπληρώσουν ?
                        ' Αν ναι έχει καλώς
                        'Αν όχι σημαίνει ότι στις ίδιες σειρές που εχω μαρκαρει (ch2<>null) πρέπει να πάρω όλη την ποσότητα
                        ' γιατί δεν θα συμπληρώσω το τιμολόγιο
                        ' αγνοώ δηλαδή το πιστόλι και πάω lifo ή fifo
                        'παράδειγμα που έτυχε

                        'PARTIDA	HME	KOD	TIMOLOGIA	TEMAXIA	ID	N1	N2	CH1	CH2	                                                    ch2	        ypol
                        '1708546	2021-05-13 00:00:00.000	001299	001319;1485-2020;2020-12-21*	7850	36802	7762	0	NULL	 7800      	7850	    
                        '1708547	2021-05-13 00:00:00.000	001299	001319;1597-2020;2021-04-06*	7750	36803	7998	230	NULL	 5850      	7750	      
                        '1708550	2021-05-14 00:00:00.000	001299	001319;1597-2020;2021-04-06*	17550	36806	7998	233	NULL	15600      17550
                        '                                                                                                               -----       -----
                        '                                                                                                               29250      33150       
                        ' αν ακολουθησω το πιστολάκι μετά δεν θα εχει αλλη παρτιδα να παρω και θα μείνει λειψό το τιμολογιο
                        Dim rr1 As New DataTable, rr2 As New DataTable, SMARKAR As Single, SUNMARK As Single
                        ExecuteSQLQuery("select  ISNULL( sum(YPOL),0) AS YP,ISNULL(SUM(CONVERT(REAL,CH2)),0)  AS MCH2 from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' and (NOT CH2 IS NULL)", rr1)
                        SMARKAR = rr1(0)(1)
                        ExecuteSQLQuery("select  ISNULL( sum(YPOL),0) AS YP from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' and ( CH2 IS NULL)", rr2)
                        SUNMARK = rr2(0)(0)

                        ' ΤΑ ΑΜΑΡΚΑΡΙΣΤΑ ΔΕΝ ΦΤΑΝΟΥΝ ΝΑ ΚΑΛΥΨΟΥΝ ΑΥΤΑ ΠΟΥ ΜΑΣ ΛΕΙΠΟΥΝ ΤΟΤΕ ΠΑΩ LIFO/FIFO ΚΑΙ ΑΓΝΟΩ ΤΟ ΜΑΡΚΑΡΙΑΣΜΑ

                        Dim rr As New DataTable
                        If SUNMARK < (mposo - SMARKAR) Then
                            ExecuteSQLQuery("select *,'' AS MCH2 from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' ORDER BY " + mSQLscan + "HME " + mFiFo, rr)


                        Else
                            ExecuteSQLQuery("select *,isnull(CH2,'') AS MCH2 from PARTIDES WHERE YPOL>0 AND KOD='" + cEID + "' ORDER BY " + mSQLscan + "HME " + mFiFo, rr)

                        End If





                        w = 111
                        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                        If NoPistotiko = -1 Then  'ΠΙΣΤΩΤΙΚΑ ΕΠΙΣΤΡΟΦΕΣ
                            ' PROSOXH ------------------------------
                            'ΣΤΑ ΠΙΣΤΩΤΙΚΑ ΣΤΟ TIMSPOL ΑΠΟΘΗΚΕΥΩ ΤΑ ΕΞΗΣ ΓΙΑ ΝΑ ΜΠΟΡΩ ΝΑ ΤΑ ΣΒΗΣΩ
                            'C1='ΠΙΣ'
                            'Ν1= ID PARTIDES
                            'N2 = ID TIMS 




                            Dim RSYN2 As New ADODB.Recordset
                            RSYN2.Open("SELECT COUNT(*) FROM SYNTAGES WHERE KOD='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            'cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
                            Dim mIDpart As Long, mIDtimagor As Long

                            If RSYN2(0).Value > 0 Then  ' EINAI ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ
                                If check_only = False Then
                                    Dim mp As Long, partides As New DataTable
                                    ExecuteSQLQuery("SELECT max(PARTIDA) FROM PARTIDES  WHERE PARTIDA LIKE '17%' ", partides)
                                    If IsDBNull(partides(0)(0)) Then
                                        mp = 1700000
                                    Else
                                        mp = partides(0)(0) + 1
                                    End If
                                    Dim HM As String = Format(Now, "MM/dd/yyyy") ' 
                                    Dim SQLQuery As String = "INSERT INTO PARTIDES (PARTIDA,HME,KOD,TIMOLOGIA,TEMAXIA,YPOL,N1) VALUES(" + Str(mp) + ",'" + HM + "','" + cEID + "','" + mAtim + "'," + Replace(Str(Math.Abs(mposo)), ",", ".") + "," + Replace(Str(Math.Abs(mposo)), ",", ".") + ",0)"
                                    GDB.Execute(SQLQuery)
                                    ExecuteSQLQuery("SELECT max(ID) FROM PARTIDES  WHERE PARTIDA LIKE '17%' ", partides)
                                    If IsDBNull(partides(0)(0)) Then
                                    Else
                                        mIDpart = partides(0)(0)
                                        mIDtimagor = 0
                                    End If

                                    'GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(POS2.Rows(LL)("YPOL")), ",", ".") + ",'" + cEID + "','" + mAtim + "')")
                                    'GDB.Execute("UPDATE TIMS SET YPOL=0 WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                                End If
                            Else  ' YLIKO 
                                If check_only = False Then


                                    sql = "insert into TIMS (HME,POSO,ATIM,KOD,PROM,YPOL) VALUES("
                                    sql = sql + "'" + mHME + "',"
                                    sql = sql + Replace(Str(Math.Abs(mposo)), ",", ".") + ","
                                    sql = sql + "'" + mAtim + "',"
                                    sql = sql + "'" + cEID + "',"
                                    sql = sql + "'" + cPel + "',"
                                    sql = sql + Replace(Str(Math.Abs(mposo)), ",", ".") + ")"
                                    ExecuteSQLQuery(sql)
                                    Dim PARTIDES As New DataTable
                                    ExecuteSQLQuery("SELECT max(ID) FROM TIMS ", PARTIDES)
                                    If IsDBNull(PARTIDES(0)(0)) Then
                                    Else
                                        mIDpart = 0
                                        mIDtimagor = PARTIDES(0)(0)
                                    End If








                                    'GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(POS2.Rows(LL)("YPOL")), ",", ".") + ",'" + cEID + "','" + mAtim + "')")
                                    'GDB.Execute("UPDATE TIMS SET YPOL=0 WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                                End If
                            End If
                            RSYN2.Close()
                            If check_only = False Then
                                sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1,N2,C1) VALUES("
                                sql = sql + "'" + mHME + "',"
                                sql = sql + Replace(Str(Math.Abs(mposo)), ",", ".") + ","
                                sql = sql + "'" + mAtim + "',"
                                sql = sql + "'" + cEID + "',"
                                sql = sql + "'" + Mid(cPel, 1, 6) + "',"
                                sql = sql + "0" + ",'" + "." + "'," + Str(mIDpart) + "," + Str(mIDtimagor) + ",'ΠΙΣ' )"
                                w = 114
                                ' r.Close()

                                'If check_only = False Then
                                GDB.Execute(sql)
                            End If


                            'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            ' χρεωστικα παραστατικά
                        Else


                            'ΔΕΝ ΒΡΗΚΑ ΚΑΘΟΛΟΥ ΠΑΡΤΙΔΕΣ ΣΕ ΑΥΤΟ ΤΟ ΕΙΔΟΣ
                            'NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN
                            'If r.EOF Then
                            If rr.Rows.Count = 0 Then
                                'ΕΙΝΑΙ ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ????
                                Dim RSYN As New ADODB.Recordset
                                RSYN.Open("SELECT COUNT(*) FROM SYNTAGES WHERE KOD='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)




                                'cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
                                If RSYN(0).Value > 0 Then  ' EINAI ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ
                                    LINES.Items.Add(";;;;ΔΕΝ ΥΠΑΡΧΕΙ ΠΑΡΤΙΔΑ ΓΙΑ ΤΟ ΕΙΔΟΣ " + cEID)
                                    MsgBox(" δεν υπαρχει καθόλου υπολοιπο παρτιδων για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                    is_ok = False
                                    Exit While
                                    'Exit Sub

                                    w = 112
                                    'sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
                                    'sql = sql + "'" + mHME + "',"
                                    'sql = sql + Kau_Aritmoy(mposo) + ","
                                    'sql = sql + "'" + mAtim + "',"
                                    'sql = sql + "'" + cEID + "',"
                                    'sql = sql + "'" + Mid(cPel, 1, 6) + "',"
                                    'sql = sql + "0" + ",'" + "." + "'," + "0" + " )"
                                    'w = 114
                                    '' r.Close()
                                    'Try
                                    '    GDB.Execute(sql)
                                    '    LINES.Items.Add("Προιόν;" + cEID + ";Τιμολ.;" + mAtim + ";" + Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10) + ";Παρτίδα;0;" + cPel)
                                    'Catch ex As Exception
                                    '    MsgBox("δεν αποθηκευτηκε το τιμολογιο " + mAtim + "  " + ex.Message)
                                    'End Try

                                    'cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
                                Else  'ΔΕΝ ΕΙΝΑΙ  ΕΤΟΙΜΟ Ή ΗΜΙΕΤΟΙΜΟ  ΑΡΑ ΕΙΝΑΙ ΠΡΩΤΗ ΥΛΗ Ή ΒΟΗΘΗΤΙΚΗ
                                    ' ΒΡΕΣ ΤΑ ΤΙΜΟΛΟΓΙΑ ΠΟΥ ΕΧΟΥΝ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ

                                    isA_YLH = 1
                                    Dim POS2 As New DataTable
                                    Dim r2 As New ADODB.Recordset
                                    r2.Open("select sum(YPOL) FROM TIMS WHERE RTRIM(LTRIM(KOD))='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                    If IsDBNull(r2(0).Value) Then
                                        ' < mposo Then
                                        MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο TIM.AΓOΡAΣ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                        is_ok = False
                                        Exit Sub
                                    Else
                                        If r2(0).Value < mposo Then
                                            MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο TIM.AΓOΡAΣ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                            is_ok = False
                                            Exit Sub
                                        End If



                                    End If
                                    r2.Close()



                                    ExecuteSQLQuery("SELECT  KOD,YPOL,ATIM,HME,ID  FROM TIMS where YPOL>0 AND RTRIM(LTRIM(KOD))='" + cEID + "' ORDER BY HME " + mFiFo, POS2)
                                    If POS2.Rows.Count = 0 Then  ' '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Else '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim LL As Integer, m_MPOSO As Single = mposo
                                        '=======///==============
                                        For LL = 0 To POS2.Rows.Count - 1
                                            'ΤΟ ΠΟΣΟ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΕΙ ΕΙΝΑΙ ΛΙΓΟΤΕΡΟ ΑΠΟ ΤΟ ΤΙΜΟΛ.ΑΓΟΡΑΣ
                                            'ΟΠΟΤΕ ΑΦΑΙΡΩ ΟΛΟ ΤΟ ΠΟΣΟ ΤΟΥ ΤΙΜΟΛ.ΠΩΛΗΣΗΣ (M_POSO)
                                            If m_MPOSO < POS2.Rows(LL)("YPOL") Then
                                                'DEBUG
                                                If check_only = False Then
                                                    GDB.Execute("UPDATE TIMS SET YPOL=YPOL-" + Replace(Str(m_MPOSO), ",", ".") + " WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                                                    GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(m_MPOSO), ",", ".") + ",'" + cEID + "','" + mAtim + "')")
                                                End If

                                                m_MPOSO = 0

                                            Else '  m_MPOSO > POS2.Rows(LL)("YPOL") Then
                                                'ΤΟ ΠΟΣΟ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΕΙ ΕΙΝΑΙ MEGALYTERO ΑΠΟ ΤΟ ΤΙΜΟΛ.ΑΓΟΡΑΣ
                                                'ΟΠΟΤΕ ΑΦΑΙΡΩ ΟΛΟ ΤΟ ΠΟΣΟ ΤΟΥ ΤΙΜΟΛ.AGORAS POS2.Rows(LL)("YPOL")
                                                m_MPOSO = m_MPOSO - POS2.Rows(LL)("YPOL")
                                                'DEBUG
                                                If check_only = False Then
                                                    GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2) VALUES (0,-3000," + POS2.Rows(LL)("ID").ToString + ",'" + mHME + "'," + Replace(Str(POS2.Rows(LL)("YPOL")), ",", ".") + ",'" + cEID + "','" + mAtim + "')")
                                                    GDB.Execute("UPDATE TIMS SET YPOL=0 WHERE ID=" + POS2.Rows(LL)("ID").ToString)
                                                End If

                                            End If
                                            If m_MPOSO = 0 Then
                                                Exit For
                                            End If
                                        Next
                                        '========///=============
                                    End If  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    '========================
                                    sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
                                    sql = sql + "'" + mHME + "',"
                                    sql = sql + Kau_Aritmoy(mposo) + ","
                                    sql = sql + "'" + mAtim + "',"
                                    sql = sql + "'" + cEID + "',"
                                    sql = sql + "'" + Mid(cPel, 1, 6) + "',"
                                    sql = sql + "0" + ",'" + "." + "'," + "0" + " )"
                                    w = 114
                                    ' r.Close()
                                    Try
                                        If check_only = False Then
                                            GDB.Execute(sql)
                                        End If

                                        LINES.Items.Add("Α ΥΛΗ " + cEID + ";Τιμολ.;" + mAtim + ";" + ";ANAΛΩΣΗ ΑΠΟ ΤΙΜΟΛΟΓΙΑ ΑΓΟΡΑΣ;" + cPel)
                                    Catch ex As Exception
                                        MsgBox("δεν αποθηκευτηκε το τιμολογιο " + mAtim + "  " + ex.Message)
                                    End Try
                                End If
                                'cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
                                w = 115

                                'End If  'NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN
                            Else ' ΒΡΗΚΑ ΠΑΡΤΙΔΕΣ

                                Dim POS2 As New DataTable
                                Dim r2 As New ADODB.Recordset
                                r2.Open("select sum(YPOL) FROM  PARTIDES WHERE RTRIM(LTRIM(KOD))='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                                If IsDBNull(r2(0).Value) Then
                                    ' < mposo Then
                                    MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο ΠΑΡΤΙΔΩΝ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                    is_ok = False
                                    Exit While
                                    'Exit Sub
                                Else
                                    If r2(0).Value < mposo Then
                                        MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο ΠΑΡΤΙΔΩΝ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                        is_ok = False
                                        Exit While
                                        'Exit Sub
                                    End If
                                End If

                                'If r2(0).Value < mposo Then
                                '    MsgBox(" δεν υπαρχει ΕΠΑΡΚΕΣ υπολοιπο ΠΑΡΤΙΔΩΝ για τον κωδικό " + cEID + Chr(13) + " το τιμολόγιο " + mAtim + " δεν θα περαστεί")
                                '    Exit Sub

                                'End If
                                r2.Close()



                                'Do While Not R.EOF
                                For LL = 0 To (rr.Rows.Count - 1)
                                    w = 12
                                    'ΤΡΑΒΑΩ ΑΠΟ ΤΗΝ ΠΡΩΤΗ ΠΑΛΙΟΤΕΡΗ ΠΑΡΤΙΔΑ
                                    If rr(LL)("YPOL") > (mposo) Then
                                        'PosoReal = (mposo)
                                        If Val(rr(LL)("MCH2")) > 0 Then
                                            PosoReal = Val(rr(LL)("mch2"))
                                        Else
                                            PosoReal = (mposo)
                                        End If




                                    Else
                                        If Val(rr(LL)("mch2")) > 0 Then
                                            PosoReal = Val(rr(LL)("mch2"))
                                        Else
                                            PosoReal = rr(LL)("YPOL")
                                        End If


                                    End If
                                    w = 13
                                    sql = "insert into TIMSPOL (HME,POSO,ATIM,KOD,PROM,PARTIDA,AFM,N1) VALUES("
                                    sql = sql + "'" + mHME + "',"
                                    sql = sql + Kau_Aritmoy(Str(PosoReal)) + ","
                                    sql = sql + "'" + mAtim + "',"
                                    sql = sql + "'" + cEID + "',"
                                    sql = sql + "'" + Mid(cPel, 1, 6) + "',"
                                    sql = sql + rr(LL)("partida").ToString + ",'" + "." + "'," + "0" + " )"
                                    w = 14
                                    If check_only = False Then
                                        GDB.Execute(sql)
                                    End If

                                    w = 15
                                    'ExecuteSQLQuery("select max(ID) FROM TIMSPOL")
                                    ' idtimols = sqlDT.Rows(0)(0).ToString



                                    'DEBUG
                                    If check_only = False Then
                                        GDB.Execute("UPDATE PARTIDES SET YPOL=YPOL-" + Kau_Aritmoy(Str(PosoReal)) + " WHERE PARTIDA=" + rr(LL)("partida").ToString)
                                        If mScan = 1 Then
                                            'TEST WORKING  UPDATE  [TECHNOPLASTIKI].[dbo].[PARTIDES] SET CH2=CONVERT(CHAR(10), CONVERT(REAL,CH2)+200)    WHERE PARTIDA='1708206'
                                            GDB.Execute("UPDATE PARTIDES SET CH2=CONVERT(CHAR(10), CONVERT(REAL,CH2) - " + Kau_Aritmoy(Str(PosoReal)) + ")  WHERE PARTIDA=" + rr(LL)("partida").ToString)

                                            'Dim FFD As New DataTable, mLastID As Integer
                                            'ExecuteSQLQuery("select max(ID) FROM TIMSPOL", FFD)
                                            ' mLastID = FFD(0)(0)
                                            ' ExecuteSQLQuery("update PALETES SET select max(ID) FROM TIMSPOL", FFD)






                                        End If
                                    End If




                                    w = 16
                                    LINES.Items.Add("Προιόν;" + cEID + ";Τιμολ.;" + ";" + mAtim + ";Παρτίδα;" + rr(LL)("partida").ToString + ";" + cPel)
                                    'Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10)




                                    If PosoReal >= Val(mposo) Then
                                        Exit For
                                    Else
                                        'αφαιρω αυτό που παρτιδοποιήθηκε
                                        mposo = mposo - PosoReal

                                    End If
                                    ' r.MoveNext()
                                    w = 17

                                Next
                                ' Loop


                                ' r.Close()

                            End If  ' ΒΡΗΚΑ ΠΑΡΤΙΔΕΣ
                            'End If  'NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN

                        End If 'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM    NoPistotiko = -1

                        w = 18
                        flagPel = 1

                        'DEBUG GDB.CommitTrans()

                        'ΕΝΗΜΕΡΩΝΩ ΤΑ TIMSANAL ME TO ID TOY TIMSPOL ΠΟΥ ΔΗΜΙΟΥΡΓΗΘΗΚΕ
                        If isA_YLH = 1 Then
                            Dim maxIDTimspol As Long
                            Dim FFD As New DataTable
                            ExecuteSQLQuery("select max(ID) FROM TIMSPOL", FFD)
                            maxIDTimspol = FFD(0)(0)
                            If check_only = False Then
                                ExecuteSQLQuery("update TIMSANAL set N2=" + Str(maxIDTimspol) + " WHERE N2=-3000 AND ID>" + Str(MaxIdTimsAnal))
                            End If

                        End If




                    Catch ex As Exception
                        MsgBox("λαθος στην σειρά " + Str(N) + Chr(13) + Err.Description)
                        'xlWorkBook.Close()
                        'xlApp.Quit()

                        'releaseObject(xlApp)
                        'releaseObject(xlWorkBook)
                        'releaseObject(xlWorkSheet)

                        If check_only = False Then
                            MsgBox("ΑΠΟΘΗΚΕΥUHKAN TA TIMOLOGIA. ΜΕΧΡΙ ΣΕΙΡΑ EXCEL " + Str(N))
                        End If
                        'DEBUG GDB.RollbackTrans()

                        Exit Sub

                    End Try
                    R32.MoveNext()

                    If R32.EOF Then
                        Exit While
                    End If

                    'Else ' ΤΕΛΕΙΩΣΕ ΤΟ ΤΙΜΟΛΟΓΙΟ ΚΑΙ ΠΑΩ ΠΑΡΑΚΑΤΩ  '============================================================================================
                    '    Exit While

                    'End If '============================================================================================
                    N = N + 1
                End While   ' λοοπ mesa sto idio timologio


                If is_ok = False Then
                    Exit Do
                End If





                If check_only = False Then
                    If mScan = 1 Then
                        GDB.Execute("UPDATE [PARTIDES] SET CH2=NULL ")
                    End If

                End If



                ' GDB.Execute("update SYNTAGES SET POSOSTO=" + Replace(Str(xlWorkSheet.Cells(N, 11).VALUE / 1), ",", ".") + " WHERE  KOD='" + cPel + "' AND KODSYNOD='" + D + "'")
            Catch ex As Exception
                MsgBox("κωδ.λαθους w=" + Str(w) + " ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + Str(N) + " " + cPel + " == " + D + "--" + ex.Message)
                'Exit Do

                MsgBox("λαθος στην σειρά " + Str(N) + Chr(13) + Err.Description)
                '  xlWorkBook.Close()
                ' xlApp.Quit()

                ' releaseObject(xlApp)
                ' releaseObject(xlWorkBook)
                ' releaseObject(xlWorkSheet)

                If check_only = False Then
                    MsgBox("ΑΠΟΘΗΚΕΥUHKAN TA TIMOLOGIA. ΜΕΧΡΙ ΣΕΙΡΑ EXCEL " + Str(N))
                End If
                'DEBUG GDB.RollbackTrans()

                Exit Sub







            End Try
            'If R32.EOF Then
            '    Exit Do
            'End If
            'R32.MoveNext()

            N = N + 1
            Me.Text = N
        Loop '  xlWorkSheet.Cells(N, 9).VALUE Is Nothing





        'ANS = MsgBox("ΝΑ ΑΠΟΘΗΚΕΥΤΟΥΝ ΟΙ ΕΙΣΑΓΩΓΕΣ ΤΙΜΟΛΟΓΙΩΝ (ΣΕΙΡΕΣ EXCEL " + Str(N) + " )  ", MsgBoxStyle.YesNo)

        'If ANS = MsgBoxResult.Yes Then
        '    GDB.CommitTrans()
        'Else
        '    GDB.RollbackTrans()
        'End If







        'xlWorkBook.Close()
        'xlApp.Quit()

        'releaseObject(xlApp)
        'releaseObject(xlWorkBook)
        'releaseObject(xlWorkSheet)

        If check_only = False Then
            'ENHMERΩΝΕΙ ΤΙΣ ΠΑΛΕΤΕΣ ΜΕ ΤΟ ΤΙΜΟΛΟΓΙΟ ΠΟΥ ΠΗΓΕ Η ΚΑΘΕ ΠΑΛΕΤΑ
            If mScan = 1 Then
                GDB.Execute("UPDATE PALETES  SET ATIMPOL=(SELECT TOP 1 ATIM FROM PALETTIM WHERE  CONVERT(INT,SUBSTRING(SCAN,13,7))=PALET)  WHERE PALET IN (SELECT CONVERT(INT,SUBSTRING(SCAN,13,7)) FROM PALETTIM) ")
            End If

            MsgBox("ΑΠΟΘΗΚΕΥUHKAN TA TIMOLOGIA. ΜΕΧΡΙ ΣΕΙΡΑ EXCEL " + Str(N))
        End If


    End Sub
End Class