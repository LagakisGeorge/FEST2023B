Public Class FrmADDSUPPLIER_ITEM


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






    Private Sub cmdcancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdcancel.Click
        Me.Close()
    End Sub

    Private Sub cmdsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsave.Click
        '       CREATE TABLE [dbo].[HOTELS](
        '[ID] [int] NOT NULL,
        '[CATEGORY] int NULL,
        '[RANK] int NULL,
        '[EMAIL] nvarchar(50),
        '[THL] NVARCHAR(30),
        '[DIE] NVARCHAR(35),
        '[NAME] [varchar](50) NULL)
        Dim SQL As String = ""
        Dim CAT2 As String = Mid(ComboCAT.Text, 1, 1)
        Dim RANK2 As String = RANK.Text





        If IsNew Then

            SQL = "insert into HOTELS (CATEGORY,RANK,EMAIL,NAME,THL,DIE) VALUES (" + CAT2 + "," + RANK2 + ",'" + txtEmail.Text + "','" + Replace(txtName.Text, "'", "`") + "','" + txtTHL.Text + "','" + txtTHL.Text + "')"

        Else
            SQL = "UPDATE HOTELS SET CATEGORY=" + CAT2 + ",RANK=" + RANK2 + ",EMAIL='" + txtEmail.Text + "',NAME='" + txtName.Text + "',THL='" + txtTHL.Text + "',DIE='" + txtDie.Text + "'  WHERE ID=" + Str(ID)


        End If



        Try
            Debug.Print(SQL)

            ExecuteSQLQuery(SQL)
        Catch ex As Exception
            MsgBox("ΔΕΝ ΚΑΤΕΧΩΡΗΘΗ " + Err.Description)
        End Try

        Me.Close()





    End Sub

    Private Sub FrmADDSUPPLIER_ITEM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtEmail.Select()
        txtEmail.Focus()
    End Sub
End Class