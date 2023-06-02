Imports CrystalDecisions.CrystalReports.Engine

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.Net.NetworkInformation
Imports System.Transactions


Public Class report

    Dim GDB As New ADODB.Connection




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        EKTYPOSI()
    End Sub
    Private Sub PROFORMESTYP(ByVal PIECESPERPALLET As Long, ByVal N2 As Integer)
        '725
        '================ προιοντα που τρώνε υλικο  μονο από προφόρμες  ==================
        Dim mORA As String = Format(Now, "HH:mm:ss")


        Dim kodPROION As String
        Dim onomaProion As String

        ListANAL.Items.Clear()

        Dim BAROS As Single
        Dim N As Integer
        If ListBox1.Items.Count = 1 Then
            N = 0
        Else
            N = ListBox1.SelectedIndex
        End If


        If N < 0 Then
            N = 0
        End If


        kodPROION = Split(ListBox1.Items(N).ToString, "*")(0)
        onomaProion = Split(ListBox1.Items(N).ToString, "*")(1)
        onomaProion = Split(onomaProion, ";")(0)


        BAROS = Val(Split(ListBox1.Items(N).ToString, ";")(1))

        ' EΠEIΔH EINAI PROFORMA ΒΑΖΩ ΒΑΡΟΣ=1 ΓΙΑ ΝΑ ΠΑΕΙ 1 ΠΡΟΣ 1
        BAROS = 1


        FillListBox("SELECT KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],STR(POSOSTO,6,3) AS [ΠΟΣΟΣΤΟ]  FROM SYNTAGES where KOD='" + kodPROION + "' ", ListBox2)

        Dim YLIKA As New DataTable

        ExecuteSQLQuery("SELECT KODSYNOD ,POSOSTO  FROM SYNTAGES where KOD='" + kodPROION + "' ", YLIKA)

        ' ΠΑΡΤΙΔΟΠΟΙΗΣΗ με minTem ετικέτες, timol() τα τιμολογια της παρτίδας


        Dim KY As Integer
        Dim YL As String
        Dim POS As Single
        Dim ApaitPosothta As Single
        Dim mT(30) As Single ' temaxia που μπορω να βγάλω από κάθε υλικό
        'χρησιμοποιώντας μόνο ένα τιμολόγιο
        Dim timol(30) As String

        Dim minTem As Long ' minimum τεμαχια που μπορουν να παραχθουν με τον ιδιο αριθμό παρτίδας
        Dim Zhtoymena_TEM As Long ' ποσα τεμαχια θέλω συνολικά
        minTem = Val(tem.Text) * Val(SYSKEYASIA.Text)
        Dim TIMOLS As String = ""

        Dim IdTims(30) As Integer
        Dim CHECKID As Long = 0 'ΣΟΥΜΑΡΕΙ ΟΛΑ ΤΑ ID ΤΩΝ ΤΙΜΟΛΟΓΙΩΝ ΓΙΑ ΝΑ ΔΕΙ ΑΝ ΑΛΛΑΞΕ ΠΑΡΤΙΔΑ

        'ypologizv ti KOMMATIA μπορει να βγαλει το τελευταίο τιμολογιο 
        ' ΔΗΛΑΔΗ Ο ΑΡΙΘΜΟΣ ΕΤΙΚΕΤΤΩΝ ΠΟΥ ΘΑ ΕΧΕΙ ΤΗΝ ΙΔΙΑ ΠΑΡΤΙΔΑ
        Dim pos20 As New DataTable
        Dim mFiFo As String
        If Mid(ComboFifo.Text, 1, 1) = 1 Then
            mFiFo = ""
        Else
            mFiFo = " desc"
        End If




        For KY = 0 To YLIKA.Rows.Count - 1
            YL = YLIKA.Rows(KY).Item(0).ToString  'ΚΩΔΙΚΟΣ ΣΥΣΤΑΤΙΚΟΥ
            POS = YLIKA.Rows(KY).Item(1)  ' ΠΟΣΟΣΤΟ ΣΥΜΜΕΤΟΧΗΣ ΣΤΗΝ ΣΥΝΤΑΓΗ


            ApaitPosothta = POS * minTem * BAROS ' P.X. 30GR   ΟΛΙΚΗ ΠΟΣΟΤΗΤΑ ΠΟΥ ΘΑ ΧΡΕΙΑΣΤΟΥΜΕ ΑΠΟ ΤΟ ΣΥΣΤΑΤΙΚΟ
            ListANAL.Items.Add(YL + " απαιτ.ποσ=" + Str(ApaitPosothta))
            If ApaitPosothta > 0 Then
                Dim pos2 As New DataTable

                ' ΒΡΕΣ (pos2)  ΤΟ ΠΡΩΤΟ ΤΙΜΟΛΟΓΙΟ ΠΟΥ ΕΧΕΙ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ
                ExecuteSQLQuery("SELECT TOP 1 KOD,YPOL,HME,ID  FROM PARTIDES where YPOL>0 AND RTRIM(LTRIM(KOD))='" + YL + "' ORDER BY HME " + mFiFo, pos2)
                If pos2.Rows.Count = 0 Then
                    mT(KY) = 0
                    ExecuteSQLQuery("SELECT ONO FROM YLIKA WHERE N1=1 and KOD='" + YL + "'")
                    MsgBox("ελλειψη απο κωδικο " + YL)  '  + "  " + sqlDT(0)(0).ToString
                Else

                    ListANAL.Items.Add("τιμ." + " υπολ= " + pos2.Rows(0).Item("ypol").ToString)

                    ' ΑΥΤΟ ΤΟ ΤΙΜΟΛΟΓΙΟ ΘΑ ΜΑΣ ΔΩΣΕΙ ΕΝΑ ΜΕΡΟΣ ΤΩΝ τεμαχιων ΠΟΥ ΘΕΛΟΥΜΕ     mt(ΚΥ) <= ΤΕΜ.ΤΕΧΤ
                    mT(KY) = Val(tem.Text) * Val(SYSKEYASIA.Text) * pos2.Rows(0).Item(1) / ApaitPosothta

                    ListANAL.Items.Add("κομματια που μπορει να δώσει  " + Str(mT(KY)))
                    If mT(KY) = 0 Then
                        ListANAL.Items.Add("***ελλειψη απο κωδικο " + YL + "  " + sqlDT(0)(0).ToString)
                        MsgBox("ελλειψη απο κωδικο " + YL + "  " + sqlDT(0)(0).ToString)
                    End If

                    timol(KY) = YL + ";" + ";" + pos2.Rows(0).Item(3).ToString + "*" ' , "dd/MM/yyyy")
                    IdTims(KY) = pos2.Rows(0).Item("ID")
                    CHECKID = CHECKID + IdTims(KY)
                    TIMOLS = TIMOLS + timol(KY)

                    ' ΕΑΝ ΤΑ ΚΟΜΜΑΤΙΑ ΠΟΥ ΔΙΝΕΙ ΕΙΝΑΙ ΛΙΓΟΤΕΡΑ ΤΟΥ 1 ΤΟΤΕ ΤΟ ΜΗΔΕΝΙΖΩ ΓΙΑ ΝΑ ΜΗΝ ΜΕ ΜΠΕΡΔΕΥΕΙ

                    'If mT(KY) < 1 Then
                    '    ExecuteSQLQuery("INSERT INTO TIMSANAL (IDPART,IDTIMS,HME,POSO,CH1) VALUES (" + "0" + "," + Str(IdTims(KY)) + ",'" + Format(Now, "MM/dd/yyyy") + "'," + toTeleia(pos2.Rows(0).Item(1).ToString) + ",'" + YL + "' )", pos20)
                    '    ExecuteSQLQuery("UPDATE TIMS SET YPOL=0  where ID=" + pos2.Rows(0).Item("id").ToString, pos20)

                    'End If






                    ' ΒΡΙΣΚΩ ΤΟΝ ΜΕΓΙΣΤΟ ΑΡΙΘΜΟ τεμαχιωνΝ ΠΟΥ ΕΧΟΥΝ ΚΟΙΝΑ ΤΙΜΟΛΟΓΙΑ 
                    ' ΓΙΑ ΝΑ ΠΑΡΟΥΝ ΤΗΝ ΙΔΙΑ ΠΑΡΤΙΔΑ
                End If

                If mT(KY) < minTem Then
                    minTem = mT(KY)
                End If



            End If
        Next

        If minTem <= 0 Then
            MsgBox(" αδυνατη η εκτυπωση λογω ελλείψεως υλικών")
            SYSKEYASIA.Text = "0"
            Exit Sub

        End If


        If minTem < Zhtoymena_TEM Then
            Me.BackColor = Color.Red
            'MsgBox("προσοχή Θα τυπωθούν μόνο " + Str(minTem) + " από τα " + Str(Zhtoymena_TEM) + " που ζήτησες. " + Chr(13) + "Ζήτησε ξανά ετικέτα")
            SYSKEYASIA.Text = Str(Val(SYSKEYASIA.Text) - minTem)
        Else
            Me.BackColor = Color.Beige
            'SYSKEYASIA.Text = Str(Zhtoymena_TEM - minTem)
            SYSKEYASIA.Text = Str(Val(SYSKEYASIA.Text) - minTem)
        End If



        ' BΡΙΣΚΩ ΤΟΝ ΑΡΙΘΜΟ ΠΑΡΤΙΔΑΣ ΠΟΥ ΘΑ ΠΑΡΕΙ
        Dim partides As New DataTable


        'ψαχνω στην ιδια μερα για ιδιο προιον αν το συνολο id (n1) ειναι ίδιο με το σύνολο id της ετικετας που τυπώνω
        ExecuteSQLQuery("SELECT TOP 1 * FROM PARTIDES WHERE KOD='" + kodPROION + "' AND DAY(HME)=DAY(GETDATE()) AND MONTH(HME)=MONTH(GETDATE()) AND YEAR(HME)=YEAR(GETDATE()) ORDER BY PARTIDA DESC ", partides)
        Dim MP As Integer
        Dim oldPartida As Long = 0 ' αν είναι >0  σημαίνει οτι συνεχιζω την ίδια παρτίδα KAI O ΑΡΙΘΜΟς ΕΙΝΑΙ Η ΠΑΡΤΙΔΑ

        Dim mN1 As Long
        If partides.Rows.Count = 0 Then
            mN1 = 0
        Else
            If IsDBNull(partides(0)("N1")) Then
                mN1 = 0
            Else
                mN1 = partides(0)("n1")
            End If

        End If


        If mN1 = CHECKID Then

            ' ειμαι στον ιδιο αριθμο παρτιδας
            MP = partides(0)("partida")
            oldPartida = partides(0)("ID")

        Else

            ExecuteSQLQuery("SELECT max(PARTIDA) FROM PARTIDES  WHERE PARTIDA LIKE '17%' ", partides)
            If IsDBNull(partides(0)(0)) Then
                MP = 1700000
            Else
                MP = partides(0)(0) + 1
            End If

        End If






        Dim HM As String = Format(Now, "MM/dd/yyyy") ' Format(Now(), "MM/DD/YYYY")
        Dim TEMAX As String = Str(minTem)

        GDB.BeginTrans()

        Dim BARC2 As String = "0000000000000000000000000"
        If minTem > PIECESPERPALLET - 1 Then
            ' BARC2 = Mid(MakeSSCC(), 2, 21)
            BARC2 = MakeSSCC()
            ' MsgBox("ΤΥΠΩΝΩ GS1")
            Dim MBARC2 As New DataTable
            ExecuteSQLQuery("INSERT INTO  PALETES  (PALET,KOD,ONO,PARTIDA,POSO,DATE) VALUES (" + Mid(BARC2, 14, 6) + ",'" + Split(ListBox1.Items(N2).ToString, "*")(0) + "','" + "" + "','" + Str(MP) + "'," + Str(minTem) + ",GETDATE() )", MBARC2)

        End If







        ' Using tran2 As New TransactionScope()
        Dim SQLQuery As String
        Try
            Dim sqlCon As New OleDbConnection(gConnect)
            If oldPartida = 0 Then ' NEA PARTIDA
                SQLQuery = "INSERT INTO PARTIDES (PARTIDA,HME,KOD,TIMOLOGIA,TEMAXIA,YPOL,N1,N2) VALUES(" + Str(MP) + ",'" + HM + "','" + kodPROION + "','" + TIMOLS + "'," + TEMAX + "," + TEMAX + "," + Str(CHECKID) + "," + Mid(BARC2, 13, 6) + ")"
            Else
                SQLQuery = "UPDATE PARTIDES SET N2=" + Mid(BARC2, 13, 6) + ", YPOL=ISNULL(YPOL, 0)+" + TEMAX + ", TEMAXIA=ISNULL(TEMAXIA, 0)+" + TEMAX + " WHERE ID=" + Str(oldPartida)
            End If

            SQLEXECUTE(SQLQuery, sqlCon)

            'ExecuteSQLQuery(SQLQuery)



            Dim mTem As New DataTable
            Dim tem22 As String
            ExecuteSQLQuery("SELECT top 1 isnull(C1,'') as C1  FROM YLIKA where KOD='" + kodPROION + "' ", mTem)
            If Mid(mTem(0)(0).ToString, 1, 3) = "100" Then
                tem22 = Replace(Str(Val(TEMAX) / 100), ",", ".")
            Else
                tem22 = TEMAX
            End If





            SQLQuery = "INSERT INTO KINEMP (PARTIDA,HME,KOD,TEMAXIA) VALUES(" + Str(MP) + ",'" + HM + "','" + kodPROION + "'," + tem22 + " )"
            GDB.Execute(SQLQuery)

            ';ExecuteSQLQuery("SELECT MAX(ID) FROM PARTIDES")
            'Dim idPart As String = sqlDT(0)(0).ToString





            ' ENHMEΡΩΝΩ ΤΑ ΤΙΜS ΜΕ ΤΙΣ ΠΟΣΟΤΗΤΕΣ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΟΥΝ

            For KY = 0 To YLIKA.Rows.Count - 1
                YL = YLIKA.Rows(KY).Item(0).ToString  'ΚΩΔΙΚΟΣ ΣΥΣΤΑΤΙΚΟΥ
                POS = YLIKA.Rows(KY).Item(1)  ' ΠΟΣΟΣΤΟ ΣΥΜΜΕΤΟΧΗΣ ΣΤΗΝ ΣΥΝΤΑΓΗ
                'ΑΠΛΟΠΟΙΗΣΗ
                ApaitPosothta = POS * BAROS * minTem  ' P.X. 30GR  TO KLASMA ΑΠΟ ΤΗΝ  ΟΛΙΚΗ ΠΟΣΟΤΗΤΑ ΠΟΥ ΘΑ ΧΡΕΙΑΣΤΟΥΜΕ ΑΠΟ ΤΟ ΣΥΣΤΑΤΙΚΟ

                ' ΒΡΕΣ ΤΟ ΠΡΩΤΟ ΤΙΜΟΛΟΓΙΟ ΠΟΥ ΕΧΕΙ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ
                SQLEXECUTE("UPDATE PARTIDES SET YPOL=YPOL-" + Str(ApaitPosothta) + " where ID=" + Str(IdTims(KY)) + " ", sqlCon)


                'IDTIMS -> ΒΑΖΩ ΤΟ ID THS PARTIDAS    N2=ΑΡΙΘΜΟΣ ΠΑΡΤΙΔΑΣ
                SQLEXECUTE("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2,TEMAX) VALUES (-1," + Str(MP) + "," + Str(IdTims(KY)) + ",'" + HM + "'," + Str(ApaitPosothta) + ",'" + YL + "','" + Format(Now, "HH:mm:ss") + "'," + Str(ApaitPosothta) + " )", sqlCon)
                ListANAL.Items.Add("ΠΑΡΤΙΔΑ :" + Str(MP) + " - ΥΛΙΚΟ:" + YL + "- ΠΟΣ:" + Str(ApaitPosothta))
            Next

            GDB.CommitTrans()

            '   tran2.Complete()
        Catch ex As TransactionAbortedException
            GDB.RollbackTrans()
            MsgBox("Error : " & ex.Message)
        End Try


        Dim TELBARCODE As String
        Dim chkdig As String = findCheckDigit("521301114" + Mid(kodPROION + Space(6), 3, 6))
        TELBARCODE = "'02521301114" + Mid(kodPROION + Space(6), 3, 4) + chkdig + "300"
        TELBARCODE = TELBARCODE + Mid(LTrim(Str(1000000 + minTem)), 3, 5) + "100" + Mid(LTrim(Str(MP)), 1, 7)


        ' 18-3-
        ' Dim TELBARCODE As String
        ' TELBARCODE = "'0205200000" + Mid(kodPROION + Space(6), 1, 6) + "300" + Mid(Str(100000 + minTem), 3, 5) + "10000" + Mid(LTrim(Str(MP)), 1, 6)



        'Dim xlApp As Excel.Application
        'Dim xlWorkBook As Excel.Workbook
        'Dim xlWorkSheet As Excel.Worksheet

        'xlApp = New Excel.ApplicationClass
        'If BARC2 <> "0000000000000000000000000" Then
        '    xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\sscc.xlsx")
        '    xlWorkSheet = xlWorkBook.Worksheets("sh1")

        '    xlWorkSheet.Cells(4, 2) = onomaProion
        '    xlWorkSheet.Cells(6, 3) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        '    xlWorkSheet.Cells(7, 2) = "ΠΑΡΤΙΔΑ/LOT :" + Str(MP)
        '    xlWorkSheet.Cells(10, 1) = TELBARCODE
        '    xlWorkSheet.Cells(5, 2) = kodPROION

        '    xlWorkSheet.Cells(15, 2) = BARC2


        'Else
        '    xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\test-label.xlsx")
        '    xlWorkSheet = xlWorkBook.Worksheets("sh1")

        '    xlWorkSheet.Cells(7, 2) = onomaProion
        '    xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        '    xlWorkSheet.Cells(9, 1) = "ΠΑΡΤΙΔΑ/LOT :" + Str(MP)
        '    xlWorkSheet.Cells(15, 1) = TELBARCODE
        '    xlWorkSheet.Cells(18, 2) = kodPROION



        'End If



        ''display the cells value B2
        ''    MsgBox(xlWorkSheet.Cells(6, 1).value)
        ''edit the cell with new value

        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)
        'xlWorkBook.Save()

        'xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        'xlWorkBook.Save()


        'xlWorkBook.Close()
        'xlApp.Quit()

        'releaseObject(xlApp)
        'releaseObject(xlWorkBook)
        'releaseObject(xlWorkSheet)













        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass

        If BARC2 <> "0000000000000000000000000" Then
            xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\sscc.xlsx")
            xlWorkSheet = xlWorkBook.Worksheets("sh1")

            xlWorkSheet.Cells(4, 2) = onomaProion
            xlWorkSheet.Cells(6, 3) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
            xlWorkSheet.Cells(7, 2) = "ΠΑΡΤΙΔΑ/LOT :" + Str(MP)
            xlWorkSheet.Cells(10, 1) = TELBARCODE
            xlWorkSheet.Cells(5, 2) = kodPROION

            xlWorkSheet.Cells(15, 2) = BARC2


        Else
            xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\test-label.xlsx")
            xlWorkSheet = xlWorkBook.Worksheets("sh1")

            xlWorkSheet.Cells(7, 2) = onomaProion
            xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
            xlWorkSheet.Cells(9, 1) = "ΠΑΡΤΙΔΑ/LOT :" + Str(MP)
            xlWorkSheet.Cells(15, 1) = TELBARCODE
            xlWorkSheet.Cells(18, 2) = kodPROION
        End If

       
        'xlApp = New Excel.ApplicationClass
        'xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\test-label.xlsx")
        'xlWorkSheet = xlWorkBook.Worksheets("sh1")
        ''display the cells value B2
        ''    MsgBox(xlWorkSheet.Cells(6, 1).value)
        ''edit the cell with new value
        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(9, 1) = "ΠΑΡΤΙΔΑ/LOT :" + Str(MP)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)
        'xlWorkBook.Save()

        xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        If minTem = Zhtoymena_TEM Then
            Me.Close()

        End If

    End Sub
    Private Sub EKTYPOSI()



        Dim N2 As Integer
        If ListBox1.Items.Count = 1 Then
            N2 = 0
        Else
            N2 = ListBox1.SelectedIndex
        End If

        Dim PIECESPERPALLET As Long

        Dim dtYliko As New DataTable
        ExecuteSQLQuery("select ISNULL(PIECESPERPALLET,0) AS NN from  YLIKA WHERE KOD='" + Split(ListBox1.Items(N2).ToString, "*")(0) + "' ", dtYliko)
        PIECESPERPALLET = dtYliko(0)(0)
        dtYliko = Nothing



        'ποιοτικος
        'ExecuteSQLQuery("INSERT INTO POIOTIKOS ( KOD ,HME  ) VALUES ('" + Split(PROION.Text, "*")(0) + "',getdate() )", R)
        Dim dtPoiot As New DataTable


        ExecuteSQLQuery("select HME from  POIOTIKOS WHERE KOD='" + Split(ListBox1.Items(N2).ToString, "*")(0) + "' ORDER BY HME DESC", dtPoiot)
        If dtPoiot.Rows.Count = 0 Then
            VARDIA.PROION.Text = ListBox1.Items(0).ToString
            VARDIA.ShowDialog()


        Else


            If Math.Abs(DateDiff(DateInterval.Hour, Now, dtPoiot(0)(0))) > 8 Then

                VARDIA.PROION.Text = ListBox1.Items(0).ToString
                VARDIA.ShowDialog()


            End If


        End If







        Dim mYl As New DataTable, mKy As Integer, mYl2 As String, mpos As Single

        ExecuteSQLQuery("SELECT KODSYNOD ,POSOSTO  FROM SYNTAGES where KOD='" + Split(ListBox1.Items(N2).ToString, "*")(0) + "' ", mYl)
        For mKy = 0 To mYl.Rows.Count - 1
            mYl2 = mYl.Rows(mKy).Item(0).ToString  'ΚΩΔΙΚΟΣ ΣΥΣΤΑΤΙΚΟΥ
            mpos = mYl.Rows(mKy).Item(1)  ' ΠΟΣΟΣΤΟ ΣΥΜΜΕΤΟΧΗΣ ΣΤΗΝ ΣΥΝΤΑΓΗ

            If mpos = 1 Then
                'ΕΛΕΓΧΟΣ ΑΝ ΤΟ ΣΥΣΤΑΤΙΚΟ ΕΙΝΑΙ ΠΑΡΑΓΟΜΕΝΟ ΑΠΟ ΕΜΑΣ (DHLADH EXEI ΣΥΝΤΑΓΗ)
                ExecuteSQLQuery("SELECT COUNT(*) FROM SYNTAGES where KOD='" + mYl2 + "' ", mYl)
                If mYl.Rows(0).Item(0) > 0 Then
                    PROFORMESTYP(PIECESPERPALLET, N2)
                    Exit Sub
                End If
            End If

        Next




        '================ προιοντα με βάρος που τρώνε υλικο  μονο από τιμολόγια  ==================


        ''ΜΟΛΙΣ ΓΙΝΕΙ Η ΠΑΡΤΙΔΑ ΝΑ ΠΑΡΕΙ ΤΟ ID THΣ ΝΕΑΣ ΠΑΡΤΙΔΑΣ
        'ExecuteSQLQuery("SELECT MAX(ID) FROM PARTIDES")
        'Dim idPart As String = sqlDT(0)(0).ToString
        'idPart = idPart + 1



        GDB.BeginTrans()

        Dim mORA As String = Format(Now, "HH:mm:ss")

        Dim kodPROION As String
        Dim onomaProion As String

        ListANAL.Items.Clear()
        '' The file system path we need to split.
        'Dim s As String = "C:\Users\Sam\Documents\Perls\Main"

        '' Split the string on the backslash character.
        'Dim parts As String() = s.Split(New Char() {"\"c})

        '' Loop through result strings with For Each.
        'Dim part As String
        'For Each part In parts
        '    Console.WriteLine(part)
        'Next

        'ExecuteSQLQuery(sql)
        Dim BAROS As Single

        Dim N As Integer
        If ListBox1.Items.Count = 1 Then
            N = 0
        Else
            N = ListBox1.SelectedIndex
        End If


        If N < 0 Then
            N = 0
        End If


        kodPROION = Split(ListBox1.Items(N).ToString, "*")(0)
        onomaProion = Split(ListBox1.Items(N).ToString, "*")(1)
        onomaProion = Split(onomaProion, ";")(0)


        BAROS = Val(Split(ListBox1.Items(N).ToString, ";")(1))
        ' KOD.Text = Split(ListBox1.Items(N).ToString, ";")(1)

        'ExecuteSQLQuery("SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2  FROM YLIKA WHERE N1=1 ORDER BY KOD ")


        FillListBox("SELECT KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],STR(POSOSTO,6,3) AS [ΠΟΣΟΣΤΟ]  FROM SYNTAGES where KOD='" + kodPROION + "' ", ListBox2)

        Dim YLIKA As New DataTable

        ExecuteSQLQuery("SELECT KODSYNOD ,POSOSTO  FROM SYNTAGES where KOD='" + kodPROION + "' ", YLIKA)

        ' ΠΑΡΤΙΔΟΠΟΙΗΣΗ με minTem ετικέτες, timol() τα τιμολογια της παρτίδας


        Dim KY As Integer
        Dim YL As String
        Dim POS As Single
        Dim ApaitPosothta As Single
        Dim mT(30) As Single ' temaxia που μπορω να βγάλω από κάθε υλικό
        'χρησιμοποιώντας μόνο ένα τιμολόγιο
        Dim timol(30) As String

        Dim minTem As Long ' minimum τεμαχια που μπορουν να παραχθουν με τον ιδιο αριθμό παρτίδας
        Dim Zhtoymena_TEM As Long ' ποσα τεμαχια θέλω συνολικά
        minTem = Val(tem.Text) * Val(SYSKEYASIA.Text)
        Dim TIMOLS As String = ""

        Dim IdTims(30) As Integer
        Dim CHECKID As Long = 0 'ΣΟΥΜΑΡΕΙ ΟΛΑ ΤΑ ID ΤΩΝ ΤΙΜΟΛΟΓΙΩΝ ΓΙΑ ΝΑ ΔΕΙ ΑΝ ΑΛΛΑΞΕ ΠΑΡΤΙΔΑ

        'ypologizv ti KOMMATIA μπορει να βγαλει το τελευταίο τιμολογιο 
        ' ΔΗΛΑΔΗ Ο ΑΡΙΘΜΟΣ ΕΤΙΚΕΤΤΩΝ ΠΟΥ ΘΑ ΕΧΕΙ ΤΗΝ ΙΔΙΑ ΠΑΡΤΙΔΑ
        Dim pos20 As New DataTable
        Dim mFiFo As String
        If Mid(ComboFifo.Text, 1, 1) = 1 Then
            mFiFo = ""
        Else
            mFiFo = " desc"
        End If


        Dim QUERY_SBHSIMO_MIKROPOSOTHTON As String = ""
        Dim QUERY_SBHSIMO_MIKROPOSOTHTON2 As String = ""



        For KY = 0 To YLIKA.Rows.Count - 1
            YL = YLIKA.Rows(KY).Item(0).ToString  'ΚΩΔΙΚΟΣ ΣΥΣΤΑΤΙΚΟΥ
            POS = YLIKA.Rows(KY).Item(1)  ' ΠΟΣΟΣΤΟ ΣΥΜΜΕΤΟΧΗΣ ΣΤΗΝ ΣΥΝΤΑΓΗ




            ApaitPosothta = POS * minTem * BAROS ' P.X. 30GR   ΟΛΙΚΗ ΠΟΣΟΤΗΤΑ ΠΟΥ ΘΑ ΧΡΕΙΑΣΤΟΥΜΕ ΑΠΟ ΤΟ ΣΥΣΤΑΤΙΚΟ
            ListANAL.Items.Add(YL + " απαιτ.ποσ=" + Str(ApaitPosothta))
            If ApaitPosothta > 0 Then
                Dim pos2 As New DataTable

                ' ΒΡΕΣ (pos2)  ΤΟ ΠΡΩΤΟ ΤΙΜΟΛΟΓΙΟ ΠΟΥ ΕΧΕΙ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ
                ExecuteSQLQuery("SELECT TOP 1 KOD,YPOL,ATIM,HME,ID  FROM TIMS where YPOL>0 AND RTRIM(LTRIM(KOD))='" + YL + "' ORDER BY HME " + mFiFo, pos2)
                If pos2.Rows.Count = 0 Then
                    mT(KY) = 0
                    ExecuteSQLQuery("SELECT ONO FROM YLIKA WHERE N1=1 and KOD='" + YL + "'")
                    MsgBox("ελλειψη απο κωδικο " + YL)  '  + "  " + sqlDT(0)(0).ToString
                Else

                    ListANAL.Items.Add("ΚΩΔ." + YL + "   τιμ." + pos2.Rows(0).Item("atim").ToString + " υπολ= " + pos2.Rows(0).Item("ypol").ToString)

                    ' ΑΥΤΟ ΤΟ ΤΙΜΟΛΟΓΙΟ ΘΑ ΜΑΣ ΔΩΣΕΙ ΕΝΑ ΜΕΡΟΣ ΤΩΝ τεμαχιων ΠΟΥ ΘΕΛΟΥΜΕ     mt(ΚΥ) <= ΤΕΜ.ΤΕΧΤ
                    mT(KY) = Val(tem.Text) * Val(SYSKEYASIA.Text) * pos2.Rows(0).Item(1) / ApaitPosothta

                    ListANAL.Items.Add("κομματια που μπορει να δώσει  " + Str(mT(KY)))
                    If mT(KY) = 0 Then
                        ListANAL.Items.Add("***ελλειψη απο κωδικο " + YL + "  " + sqlDT(0)(0).ToString)
                        MsgBox("ελλειψη απο κωδικο " + YL + "  " + sqlDT(0)(0).ToString)
                    End If

                    timol(KY) = YL + ";" + pos2.Rows(0).Item("atim").ToString + ";" + pos2.Rows(0).Item(3).ToString + "*" ' , "dd/MM/yyyy")
                    IdTims(KY) = pos2.Rows(0).Item("ID")
                    CHECKID = CHECKID + IdTims(KY)
                    TIMOLS = TIMOLS + timol(KY)

                    ' ΕΑΝ ΤΑ ΚΟΜΜΑΤΙΑ ΠΟΥ ΔΙΝΕΙ ΕΙΝΑΙ ΛΙΓΟΤΕΡΑ ΤΟΥ 1 ΤΟΤΕ ΤΟ ΜΗΔΕΝΙΖΩ ΓΙΑ ΝΑ ΜΗΝ ΜΕ ΜΠΕΡΔΕΥΕΙ

                    If mT(KY) < 1 Then
                        QUERY_SBHSIMO_MIKROPOSOTHTON = "INSERT INTO TIMSANAL (IDPART,IDTIMS,HME,POSO,CH1,CH2) VALUES (" + "0" + "," + Str(IdTims(KY)) + ",'" + Format(Now, "MM/dd/yyyy") + "'," + toTeleia(pos2.Rows(0).Item(1).ToString) + ",'" + YL + "','" + mORA + "')"
                        GDB.Execute(QUERY_SBHSIMO_MIKROPOSOTHTON)
                        '  ExecuteSQLQuery("INSERT INTO TIMSANAL (IDPART,IDTIMS,HME,POSO,CH1,CH2) VALUES (" + "0" + "," + Str(IdTims(KY)) + ",'" + Format(Now, "MM/dd/yyyy") + "'," + toTeleia(pos2.Rows(0).Item(1).ToString) + ",'" + YL + "',STR(DATEPART(HOUR,GETDATE())  )+':'+LTRIM(STR(DATEPART(HOUR,GETDATE())  ))+':'+LTRIM(STR(DATEPART(SECOND,GETDATE())  )))", pos20)

                        QUERY_SBHSIMO_MIKROPOSOTHTON2 = "UPDATE TIMS SET YPOL=0  where ID=" + pos2.Rows(0).Item("id").ToString
                        GDB.Execute(QUERY_SBHSIMO_MIKROPOSOTHTON2) ' , pos20)
                        ' ExecuteSQLQuery("UPDATE TIMS SET YPOL=0  where ID=" + pos2.Rows(0).Item("id").ToString, pos20)
                    End If






                    ' ΒΡΙΣΚΩ ΤΟΝ ΜΕΓΙΣΤΟ ΑΡΙΘΜΟ τεμαχιωνΝ ΠΟΥ ΕΧΟΥΝ ΚΟΙΝΑ ΤΙΜΟΛΟΓΙΑ 
                    ' ΓΙΑ ΝΑ ΠΑΡΟΥΝ ΤΗΝ ΙΔΙΑ ΠΑΡΤΙΔΑ
                End If

                If mT(KY) < minTem Then
                    minTem = mT(KY)
                End If



            End If
        Next

        If minTem <= 0 Then
            MsgBox(" αδυνατη η εκτυπωση λογω ελλείψεως υλικών")
            GDB.RollbackTrans()



            'ΠΡΕΠΕΙ ΝΑ ΣΒΗΣΩ ΤΙΣ ΜΙΚΡΟΠΟΣΟΤΗΤΕΣ ΓΙΑ ΝΑ ΜΠΟΡΕΙ ΝΑ ΠΑΡΕΙ ΤΟ ΕΠΟΜΕΝΟ ΤΙΜΟΛΟΓΙΟ
            '(ΤΙΣ ΕΧΩ ΣΒΗΣΕΙ ΠΑΡΑΠΑΝΩ ΑΛΛΑ ΑΝ ΓΙΝΕΙ ΑΚΥΡΩΣΗ ΚΑΝΕΙ ROLLBACK ΟΠΟΤΕ ΠΡΕΠΕΙ ΝΑ ΤΟ ΞΑΝΑΣΒΗΣΩ)
            If Len(QUERY_SBHSIMO_MIKROPOSOTHTON) > 0 Then
                GDB.Execute(QUERY_SBHSIMO_MIKROPOSOTHTON)
            End If

            If Len(QUERY_SBHSIMO_MIKROPOSOTHTON2) > 0 Then
                GDB.Execute(QUERY_SBHSIMO_MIKROPOSOTHTON2)
            End If

            Exit Sub

        End If


        If minTem < Zhtoymena_TEM Then
            Me.BackColor = Color.Red
            'MsgBox("προσοχή Θα τυπωθούν μόνο " + Str(minTem) + " από τα " + Str(Zhtoymena_TEM) + " που ζήτησες. " + Chr(13) + "Ζήτησε ξανά ετικέτα")
            SYSKEYASIA.Text = Str(Val(SYSKEYASIA.Text) - minTem)
        Else
            SYSKEYASIA.Text = Str(Val(SYSKEYASIA.Text) - minTem)
            Me.BackColor = Color.Beige
        End If



        ' BΡΙΣΚΩ ΤΟΝ ΑΡΙΘΜΟ ΠΑΡΤΙΔΑΣ ΠΟΥ ΘΑ ΠΑΡΕΙ
        Dim partides As New DataTable


        'ψαχνω στην ιδια μερα για ιδιο προιον αν το συνολο id (n1) ειναι ίδιο με το σύνολο id της ετικετας που τυπώνω
        ExecuteSQLQuery("SELECT TOP 1 * FROM PARTIDES WHERE KOD='" + kodPROION + "' AND DAY(HME)=DAY(GETDATE()) AND MONTH(HME)=MONTH(GETDATE()) AND YEAR(HME)=YEAR(GETDATE()) ORDER BY PARTIDA DESC ", partides)
        Dim MP As Integer
        Dim oldPartida As Long = 0 ' αν είναι >0  σημαίνει οτι συνεχιζω την ίδια παρτίδα KAI O ΑΡΙΘΜΟς ΕΙΝΑΙ Η ΠΑΡΤΙΔΑ

        Dim mN1 As Long
        If partides.Rows.Count = 0 Then
            mN1 = 0
        Else
            If IsDBNull(partides(0)("N1")) Then
                mN1 = 0
            Else
                mN1 = partides(0)("n1")
            End If

        End If


        If mN1 = CHECKID Then
            ' ειμαι στον ιδιο αριθμο παρτιδας
            MP = partides(0)("partida")
            oldPartida = partides(0)("ID")
        Else
            ExecuteSQLQuery("SELECT max(PARTIDA) FROM PARTIDES WHERE PARTIDA LIKE '17%' ", partides)
            If IsDBNull(partides(0)(0)) Then
                MP = 1700000
            Else
                MP = partides(0)(0) + 1
            End If

        End If






        Dim HM As String = Format(Now, "MM/dd/yyyy") ' Format(Now(), "MM/DD/YYYY")
        Dim TEMAX As String = Str(minTem)
        'Select [PARTIDA],[HME],[KOD],[TIMOLOGIA,[TEMAXIA],[ID],[N1],[N2],[CH1],[CH2] FROM [dbo].[PARTIDES]

        '2990        2990-1

        Dim BARC2 As String = "0000000000000000000000000"
        If minTem > PIECESPERPALLET - 1 Then
            ' BARC2 = Mid(MakeSSCC(), 2, 22)
            BARC2 = MakeSSCC()

            ' MsgBox("ΤΥΠΩΝΩ GS1")
            Dim MBARC2 As New DataTable
            ExecuteSQLQuery("INSERT INTO  PALETES  (PALET,KOD,ONO,PARTIDA,POSO,DATE) VALUES (" + Mid(BARC2, 14, 6) + ",'" + Split(ListBox1.Items(N2).ToString, "*")(0) + "','" + "" + "','" + Str(MP) + "'," + Str(minTem) + ",GETDATE() )", MBARC2)

        End If








        ' ENHMEΡΩΝΩ ΤΑ ΤΙΜS ΜΕ ΤΙΣ ΠΟΣΟΤΗΤΕΣ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΟΥΝ

        'For KY = 0 To YLIKA.Rows.Count - 1
        '    YL = YLIKA.Rows(KY).Item(0).ToString  'ΚΩΔΙΚΟΣ ΣΥΣΤΑΤΙΚΟΥ
        '    POS = YLIKA.Rows(KY).Item(1)  ' ΠΟΣΟΣΤΟ ΣΥΜΜΕΤΟΧΗΣ ΣΤΗΝ ΣΥΝΤΑΓΗ
        '    'ApaitPosothta = POS * Val(tem.Text) * BAROS * (minTem / Val(tem.Text)) ' P.X. 30GR  TO KLASMA ΑΠΟ ΤΗΝ  ΟΛΙΚΗ ΠΟΣΟΤΗΤΑ ΠΟΥ ΘΑ ΧΡΕΙΑΣΤΟΥΜΕ ΑΠΟ ΤΟ ΣΥΣΤΑΤΙΚΟ

        '    'ΑΠΛΟΠΟΙΗΣΗ
        '    ApaitPosothta = POS * BAROS * minTem  ' P.X. 30GR  TO KLASMA ΑΠΟ ΤΗΝ  ΟΛΙΚΗ ΠΟΣΟΤΗΤΑ ΠΟΥ ΘΑ ΧΡΕΙΑΣΤΟΥΜΕ ΑΠΟ ΤΟ ΣΥΣΤΑΤΙΚΟ
        '    Dim pos2 As New DataTable
        '    ' ΒΡΕΣ ΤΟ ΠΡΩΤΟ ΤΙΜΟΛΟΓΙΟ ΠΟΥ ΕΧΕΙ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ
        '    ExecuteSQLQuery("UPDATE TIMS SET YPOL=YPOL-" + Replace(Str(ApaitPosothta), ",", ".") + " where ID=" + Str(IdTims(KY)) + " ", pos2)

        'Next




        ' Using tran2 As New TransactionScope()
        Dim SQLQuery As String
        Dim SQLQuery22 As String
        Try
            Dim sqlCon As New OleDbConnection(gConnect)
            If oldPartida = 0 Then ' NEA PARTIDA
                SQLQuery22 = "INSERT INTO PARTIDES (PARTIDA,HME,KOD,TIMOLOGIA,TEMAXIA,YPOL,N1,N2) VALUES(" + Str(MP) + ",'" + HM + "','" + kodPROION + "','" + TIMOLS + "'," + TEMAX + "," + TEMAX + "," + Str(CHECKID) + "," + Mid(BARC2, 13, 6) + ")"
            Else
                SQLQuery22 = "UPDATE PARTIDES SET N2=" + Mid(BARC2, 13, 6) + ",YPOL=ISNULL(YPOL, 0)+" + TEMAX + ", TEMAXIA=ISNULL(TEMAXIA, 0)+" + TEMAX + " WHERE ID=" + Str(oldPartida)
            End If

            GDB.Execute(SQLQuery22) ', sqlCon)




            '----

            Dim mTem As New DataTable
            Dim tem22 As String
            ExecuteSQLQuery("SELECT top 1 isnull(C1,'') as C1  FROM YLIKA where KOD='" + kodPROION + "' ", mTem)
            If Mid(mTem(0)(0).ToString, 1, 3) = "100" Then
                tem22 = Replace(Str(Val(TEMAX) / 100), ",", ".")
            Else
                tem22 = TEMAX
            End If

            SQLQuery22 = "INSERT INTO KINEMP (PARTIDA,HME,KOD,TEMAXIA) VALUES(" + Str(MP) + ",'" + HM + "','" + kodPROION + "'," + tem22 + " )"
            GDB.Execute(SQLQuery22)



            'ExecuteSQLQuery(SQLQuery)


            'ExecuteSQLQuery("SELECT MAX(ID) FROM PARTIDES")
            'Dim idPart As String = sqlDT(0)(0).ToString

            'ExecuteSQLQuery("SELECT ID FROM PARTIDES WHERE PARTIDA=" + Str(MP))
            Dim idPart As String = "0" ' sqlDT(0)(0).ToString



            ' ENHMEΡΩΝΩ ΤΑ ΤΙΜS ΜΕ ΤΙΣ ΠΟΣΟΤΗΤΕΣ ΠΟΥ ΘΑ ΑΦΑΙΡΕΘΟΥΝ

            For KY = 0 To YLIKA.Rows.Count - 1
                YL = YLIKA.Rows(KY).Item(0).ToString  'ΚΩΔΙΚΟΣ ΣΥΣΤΑΤΙΚΟΥ
                POS = YLIKA.Rows(KY).Item(1)  ' ΠΟΣΟΣΤΟ ΣΥΜΜΕΤΟΧΗΣ ΣΤΗΝ ΣΥΝΤΑΓΗ
                'ΑΠΛΟΠΟΙΗΣΗ
                ApaitPosothta = POS * BAROS * minTem  ' P.X. 30GR  TO KLASMA ΑΠΟ ΤΗΝ  ΟΛΙΚΗ ΠΟΣΟΤΗΤΑ ΠΟΥ ΘΑ ΧΡΕΙΑΣΤΟΥΜΕ ΑΠΟ ΤΟ ΣΥΣΤΑΤΙΚΟ

                ' ΒΡΕΣ ΤΟ ΠΡΩΤΟ ΤΙΜΟΛΟΓΙΟ ΠΟΥ ΕΧΕΙ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ

                GDB.Execute("UPDATE TIMS SET YPOL=YPOL-" + Str(ApaitPosothta) + " where ID=" + Str(IdTims(KY)) + " ")
                GDB.Execute("INSERT INTO TIMSANAL (IDPART,N2,IDTIMS,HME,POSO,CH1,CH2,TEMAX) VALUES (0," + Str(MP) + "," + Str(IdTims(KY)) + ",'" + HM + "'," + Str(ApaitPosothta) + ",'" + YL + "','" + mORA + "'," + TEMAX + ")")
                ListANAL.Items.Add("ΠΑΡΤΙΔΑ :" + Str(MP) + " - ΥΛΙΚΟ:" + YL + "- ΠΟΣ:" + Str(ApaitPosothta))
            Next



            '   tran2.Complete()
        Catch ex As TransactionAbortedException
            MsgBox("Error : " & ex.Message)
        End Try
        'End Using

        'CREATE TABLE [dbo].[TIMSANAL](
        '	[ID] [int] IDENTITY(1,1) NOT NULL,
        '	[IDTIMS] [int] NOT NULL,
        '	[POSO] [real] NULL,
        '	[IDPART] [int] NOT NULL,
        '	[HME] [datetime] NULL,
        '	[N1] [int] NULL,
        '	[N2] [int] NULL,
        '	[CH1] [nvarchar](50) NULL,
        '	[CH2] [nvarchar](50) NULL
        ') ON [PRIMARY]





        'Dim connString As String = ConfigurationManager.ConnectionStrings("db").ConnectionString
        'Using conn = New SqlConnection(connString)
        '    conn.Open()
        '    Using tran As IDbTransaction = conn.BeginTransaction()
        '        Try
        '            ' transactional code...
        '            Using cmd As SqlCommand = conn.CreateCommand()
        '                cmd.CommandText = "INSERT INTO Data(Code) VALUES('A-100');"
        '                cmd.Transaction = TryCast(tran, SqlTransaction)
        '                cmd.ExecuteNonQuery()
        '            End Using
        '            tran.Commit()
        '        Catch ex As Exception
        '            tran.Rollback()
        '            Throw
        '        End Try
        '    End Using
        'End Using





        '    tem.Text = tem.Text - minTem






        ' Exit Sub
        Dim TELBARCODE As String
        'ok eos 31-12-2020  TELBARCODE = "'0205200000" + Mid(kodPROION + Space(6), 1, 6) + "300" + Mid(Str(100000 + minTem), 3, 5) + "10000" + Mid(LTrim(Str(MP)), 1, 7)

        Dim chkdig As String = findCheckDigit("521301114" + Mid(kodPROION + Space(6), 3, 6))
        TELBARCODE = "'02521301114" + Mid(kodPROION + Space(6), 3, 4) + chkdig + "300"
        TELBARCODE = TELBARCODE + Mid(LTrim(Str(1000000 + minTem)), 3, 5) + "100" + Mid(LTrim(Str(MP)), 1, 7)




        '"UPDATE MEM SET MEMO='" + onomaProion + "',  102990
        'TPSKETO=" + Str(minTem) + ",HME=GETDATE(),
        'PAR2='" + kodPROION + "',PAR1='" + TELBARCODE + "'")






        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        If BARC2 <> "0000000000000000000000000" Then
            xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\sscc.xlsx")
            xlWorkSheet = xlWorkBook.Worksheets("sh1")

            xlWorkSheet.Cells(4, 2) = onomaProion
            xlWorkSheet.Cells(6, 3) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
            xlWorkSheet.Cells(7, 2) = "ΠΑΡΤΙΔΑ/LOT :" + Str(MP)
            xlWorkSheet.Cells(10, 1) = TELBARCODE
            xlWorkSheet.Cells(5, 2) = kodPROION

            xlWorkSheet.Cells(15, 2) = BARC2


        Else
            xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\test-label.xlsx")
            xlWorkSheet = xlWorkBook.Worksheets("sh1")

            xlWorkSheet.Cells(7, 2) = onomaProion
            xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
            xlWorkSheet.Cells(9, 1) = "ΠΑΡΤΙΔΑ/LOT :" + Str(MP)
            xlWorkSheet.Cells(15, 1) = TELBARCODE
            xlWorkSheet.Cells(18, 2) = kodPROION



        End If



        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value
        
        'Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)
        xlWorkBook.Save()

        xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
























        If minTem = Zhtoymena_TEM Then
            Me.Close()

        End If




        GDB.CommitTrans()



















        Exit Sub


        '        ExecuteSQLQuery("UPDATE MEM SET MEMO='" + onomaProion + "',TPSKETO=" + Str(minTem) + ",HME=GETDATE(),PAR2='" + kodPROION + "',PAR1='" + TELBARCODE + "'")
        ' AxCrystalReport1.Action = 1

        ' Exit Sub

        'Dim cryRpt As New ReportDocument
        'Dim dscmd As New SqlDataAdapter(Sql, cnn)
        '       Dim ds As New DataTable


        'cryRpt.Load("c:\mercvb\CrystalReport1.rpt")
        ' cryRpt.SetDatabaseLogon("sa", "p@ssw0rd")    'SetDatabaseLogin(id, pw)
        'cry()
        'cryRpt.DataSourceConnections = gConnect


        'dscmd.Fill(ds, "Product")
        'cryRpt.RecordSelectionFormula = "{MEM.ID}=1" '+ mhnas.Text

        'cryRpt.PrintToPrinter(1, True, 0, 0) 'This prints one copy of all pages to the default printer, and collates them

        '      Exit Sub


        'ds = ExecuteSQLQuery("SELECT * FROM PATIENTDETAIL D INNER JOIN PATIENTS P ON P.ID=D.ID WHERE D.MHNAS='201505'")

        'cryRpt.SetDataSource(ds)






        ' CrystalReportViewer1.ReportSource = cryRpt
        ' CrystalReportViewer1.Refresh()


    End Sub


    Private Sub SQLEXECUTE(ByVal QUERY As String, ByVal sqlCon As OleDbConnection)

        Dim sqlDA As New OleDbDataAdapter(QUERY, sqlCon)
        Dim sqlCB As New OleDbCommandBuilder(sqlDA)
        sqlDT.Reset() ' refresh 
        sqlDA.Fill(sqlDT)


    End Sub

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

    '    Dim cryRpt As New ReportDocument



    '    'Dim dscmd As New SqlDataAdapter(Sql, cnn)
    '    Dim ds As New DataTable


    '    cryRpt.Load("CrystalReporteis.rpt")
    '    cryRpt.SetDatabaseLogon("sa", "12345678")    'SetDatabaseLogin(id, pw)
    '    ' cryRpt.DataSourceConnections = gConnect


    '    'dscmd.Fill(ds, "Product")
    '    cryRpt.RecordSelectionFormula = "{PATIENTS.MHNAS}=" + mhnas.Text



    '    'ds = ExecuteSQLQuery("SELECT * FROM PATIENTDETAIL D INNER JOIN PATIENTS P ON P.ID=D.ID WHERE D.MHNAS='201505'")

    '    'cryRpt.SetDataSource(ds)





    '    ' cryRpt.RecordSelectionFormula = "MHNAS=201506"
    '    CrystalReportViewer1.ReportSource = cryRpt
    '    CrystalReportViewer1.Refresh()

    '    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Len(EIDKOD.Text) = 0 Then
            FillListBox("SELECT KOD+'*'+ONO AS [ΕΙΔΟΣ],STR(BAROS,6,3) AS ΒΑΡ  FROM YLIKA WHERE (N1=2 OR N1=4) AND ONO LIKE '" + EIDOS.Text + "%'", ListBox1)
        Else
            FillListBox("SELECT KOD+'*'+ONO AS [ΕΙΔΟΣ],STR(BAROS,6,3) AS ΒΑΡ  FROM YLIKA WHERE (N1=2 OR N1=4) AND KOD LIKE '" + EIDKOD.Text + "%'", ListBox1)

        End If

    End Sub

    Private Sub BarcodeInput_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BarcodeInput.KeyUp
        Dim mkod As String
        Dim mSysk As String
        Dim sql2 As DataTable
        '000011 300 02860
        '6 ψηφ               3ψηφ  5ψηφια
        ' κωδ.προιοντος()    300   000100 συσκευασία
        If e.KeyCode = 13 Then
            mkod = Mid(BarcodeInput.Text, 1, 6)
            mSysk = Mid(BarcodeInput.Text, 10, 5)
            tem.Text = 1
            SYSKEYASIA.Text = mSysk
            'ExecuteSQLQuery("select KOD+'*'+ONO AS [ΕΙΔΟΣ],STR(BAROS,6,3) AS ΒΑΡ  FROM YLIKA WHERE N1=4 AND ONO LIKE '" + EIDOS.Text + "%'", sql2)
            FillListBox("SELECT KOD+'*'+ONO+'*'+STR(N1) AS [ΕΙΔΟΣ],STR(BAROS,6,4) AS ΒΑΡ,N1  FROM YLIKA WHERE (N1=2 OR N1=4) AND KOD='" + mkod + "'", ListBox1)
            If ListBox1.Items.Count = 0 Then
                MsgBox("ΔΕΝ ΥΠΑΡΧΕΙ")
                Exit Sub

            End If



            'If Val(Split(ListBox1.Items(N).ToString, "*")(2)) = 2 Then
            '    'ΕΙΝΑΙ ΕΜΠΟΡΕΥΜΑ ΚΑΙ ΤΟ ΑΝΟΙΓΩ ΑΥΤΟΜΑΤΑ ΣΥΝΤΑΓΗ 1->1
            'Else
            'End If
            Dim N As Integer = ListBox1.SelectedIndex
            Dim kodPROION As String = Split(ListBox1.Items(0).ToString, "*")(0)
            Dim YLIKA As New DataTable
            ExecuteSQLQuery("SELECT KODSYNOD ,POSOSTO  FROM SYNTAGES where KOD='" + kodPROION + "' ", YLIKA)

            If YLIKA.Rows.Count = 0 Then

                Dim ANS As Integer = MsgBox("ΔΕΝ ΥΠΑΡΧΕΙ ΣΥΝΤΑΓΗ. ΝΑ ΔΗΜΙΟΥΡΓΗΘΕΙ ΣΥΝΤΑΓΗ 1 ΠΡΟΣ 1", MsgBoxStyle.YesNo)
                If ANS = vbYes Then

                    ExecuteSQLQuery("INSERT INTO SYNTAGES (KODSYNOD,POSOSTO,KOD) VALUES ('" + kodPROION + "',1,'" + kodPROION + "')")

                    ExecuteSQLQuery("UPDATE YLIKA SET BAROS=1 WHERE KOD='" + kodPROION + "'")

                    ListBox1.Items.Clear()

                    FillListBox("SELECT KOD+'*'+ONO+'*'+STR(N1) AS [ΕΙΔΟΣ],STR(BAROS,6,3) AS ΒΑΡ,N1  FROM YLIKA WHERE (N1=2 OR N1=4) AND KOD='" + mkod + "'", ListBox1)

                Else
                    MsgBox("ΑΚΥΡΩΣΗ ΕΚΤΥΠΩΣΗΣ")
                    Exit Sub
                End If

            End If

            EKTYPOSI()
            Dim N3 As Integer = 1
            Do While Val(SYSKEYASIA.Text) > 0
                EKTYPOSI()
                N3 = N3 + 1

                If Val(SYSKEYASIA.Text) > 0 Then
                    Try
                        MsgBox("ΠΡΟΣΟΧΗ ΘΑ ΒΓΕΙ ΚΑΙ ΑΛΛΟ ΤΑΜΠΕΛΑΚΙ ΓΙΑ AΑΛΛΑ " + Chr(13) + SYSKEYASIA.Text + " TEMAXIA")
                    Catch ex As Exception

                    End Try

                End If


                If N3 > 10 Then
                    Exit Do
                End If
            Loop
            BarcodeInput.Text = ""


        End If



    End Sub

    Private Sub BarcodeInput_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarcodeInput.TextChanged

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\test-label.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sh1")
        'display the cells value B2
        ' MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value
        '  xlWorkSheet.Cells(1, 9) = "http://vb.net-informations.com"
        ' xlWorkSheet.PrintOutEx()


        xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=False)


        xlWorkBook.Save()


        xlWorkBook.Close()



        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)















    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strFile As String = "C:\mercvb\test-label.xlsx"
        Dim objProcess As New System.Diagnostics.ProcessStartInfo

        Dim strPrinterName As String = "ZDesigner ZT220-200dpi ZPL"

        With objProcess
            .FileName = strFile
            .WindowStyle = ProcessWindowStyle.Hidden
            .Verb = "printTo"
            .Arguments = """" & strPrinterName & """"
            .CreateNoWindow = True
            .UseShellExecute = True
        End With
        Try
            System.Diagnostics.Process.Start(objProcess)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '        Microsoft.Office.Interop.Excel.Application xlexcel;
        'Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        'Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        'Microsoft.Office.Interop.Excel.Range xlRange;
        'object misValue = System.Reflection.Missing.Value;
        'This goes in the later part of the code.

        '// Get the current printer
        'string Defprinter = null;
        'Defprinter = xlexcel.ActivePrinter;

        '// Set the printer to Microsoft XPS Document Writer
        'xlexcel.ActivePrinter = "Microsoft XPS Document Writer on Ne01:";

        '// Setup our sheet
        'var _with1 = xlWorkSheet.PageSetup;
        '// A4 papersize
        '_with1.PaperSize = Excel.XlPaperSize.xlPaperA4;
        '// Landscape orientation
        '_with1.Orientation = Excel.XlPageOrientation.xlLandscape;
        '// Fit Sheet on One Page 
        '_with1.FitToPagesWide = 1;
        '_with1.FitToPagesTall = 1;
        '// Normal Margins
        '_with1.LeftMargin = xlexcel.InchesToPoints(0.7);
        '_with1.RightMargin = xlexcel.InchesToPoints(0.7);
        '_with1.TopMargin = xlexcel.InchesToPoints(0.75);
        '_with1.BottomMargin = xlexcel.InchesToPoints(0.75);
        '_with1.HeaderMargin = xlexcel.InchesToPoints(0.3);
        '_with1.FooterMargin = xlexcel.InchesToPoints(0.3);

        '// Print the range
        'xlRange.PrintOutEx(misValue, misValue, misValue, misValue, 
        'misValue, misValue, misValue, misValue);

        '// Set printer back to what it was
        'xlexcel.ActivePrinter = Defprinter;
    End Sub

    Private Sub report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboFifo.Text = ComboFifo.Items(0).ToString
        GDB.Open(gConnect)
    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        VARDIA.PROION.Text = ListBox1.Items(0).ToString


        VARDIA.ShowDialog()

    End Sub

    Private Sub sketh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sketh.Click

        Dim BARC2 As String = ""
        ' If minTem > PIECESPERPALLET - 1 Then
        BARC2 = MakeSSCC()
        ' MsgBox("ΤΥΠΩΝΩ GS1")

        ' End If

        Dim BB As New DataTable

        ExecuteSQLQuery("insert into PALETES (PALET,DATE) VALUES (" + Mid(BARC2, 13, 7) + ",GETDATE() )", BB)


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\SKETOSSCC.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sh1")
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value
        xlWorkSheet.Cells(9, 4) = Mid(BARC2, 13, 7)

        xlWorkSheet.Cells(25, 1) = "'" + BARC2

        xlWorkBook.Save()

        xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        ' If minTem = Zhtoymena_TEM Then
        Me.Close()

        'End If

    End Sub
End Class