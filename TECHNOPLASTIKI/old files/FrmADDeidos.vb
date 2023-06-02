Public Class FrmAddEidos
    Dim stockID As Integer
    Dim hOldID As Integer
    Dim M_ID As Long = 0
    Dim mIsNew As Boolean = False

    Public Property IsNew() As Integer
        Get
            Return mIsNew
        End Get
        Set(ByVal Value As Integer)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            mIsNew = Value
            'End If
        End Set
    End Property

    Public Property ID() As Integer
        Get
            Return m_ID
        End Get
        Set(ByVal Value As Integer)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            m_ID = Value
            'End If
        End Set
    End Property




    Private Sub cmdsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsave.Click
        Dim SQL As String
        Dim mMON As String = Str(Val(MON.Text))

        Dim mb As String = BAROS.Text
        mb = Str(Val(mb))
        If Len(KOD.Text) = 0 Then
            MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΚΩΔΙΚΟ")
            Exit Sub
        End If
        If Len(ONO.Text) = 0 Then
            MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΠΕΡΙΓΡΑΦΗ")
            Exit Sub
        End If
        If Len(MON.Text) = 0 Then
            MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ")
            Exit Sub
        End If


        
        Dim mkod As String = KOD.Text
        Dim mono As String = ONO.Text
        Dim m_mon As String = MON.Text

        Dim mBaros As String = BAROS.Text

        If Val(BAROS.Text) = 0 Then
            mBaros = "0"
        End If




        If IsNew Then

            SQL = "insert into YLIKA (N1,KOD,ONO,C1,BAROS) VALUES (" + n1.Text + ",'" + KOD.Text + "','" + Replace(ONO.Text, "'", "`") + "','" + MON.Text + "'," + mb + ")"

        Else
            SQL = "UPDATE YLIKA SET KOD='" + mkod + "',ONO='" + mono + "',C1='" + m_mon + "',BAROS=" + mBaros + " WHERE ID=" + Str(ID)


        End If



        Try
            ExecuteSQLQuery(SQL)
        Catch ex As Exception
            MsgBox("ΔΕΝ ΚΑΤΕΧΩΡΗΘΗ " + Err.Description)
        End Try

        Me.Close()

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub FrmAddSupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim SQL As String
        Dim mMON As String = Str(Val(MON.Text))

        Dim mb As String = BAROS.Text
        mb = Str(Val(mb))
        If Len(KOD.Text) = 0 Then
            MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΚΩΔΙΚΟ")
            Exit Sub
        End If
        If Len(ONO.Text) = 0 Then
            MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΠΕΡΙΓΡΑΦΗ")
            Exit Sub
        End If
        If Len(MON.Text) = 0 Then
            MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ")
            Exit Sub
        End If

    End Sub
End Class