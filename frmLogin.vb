#Region "Imports"
Imports System.Data.SqlClient
Imports xAlimento
Imports GGeral
Imports System.ComponentModel
Imports System.Data.OleDb
Imports Negocios
Imports TratamentoDeErros
Imports System.Drawing.Printing
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Reflection
Imports System.Security
Imports System.Security.Principal.WindowsIdentity

    Private Sub Login()
        Dim AcessoPermitido As Boolean

        If txtUsuario.Text = "" Then
            MsgBox("Digite um nome de usuário", MsgBoxStyle.Critical, "ATENÇÃO!")


        ElseIf txtSenha.Text = "" Then
            MsgBox("Digite uma senha!", MsgBoxStyle.Critical, "ATENÇÃO!")
            txtSenha.Focus()

        Else

            Try
                Using cn = New SqlConnection(StrConnection)
                    cn.Open()

                    Using cmd = New SqlCommand("select nome, senha from Funcionario where nome =@nome and senha =@senha", cn)

                        cmd.Parameters.AddWithValue("@nome", txtUsuario.Text)
                        cmd.Parameters.AddWithValue("@senha", txtSenha.Text)

                        Using dr = cmd.ExecuteReader()
                            If dr.HasRows Then
                                If dr.Read() Then
                                    AcessoPermitido = True
                                End If
                            End If
                        End Using
                    End Using
                End Using

            Catch ex As Exception
                MsgBox("Falha ao conectar!" & vbNewLine & ex.Message, vbCritical)
            End Try

            If AcessoPermitido Then
                frmMDIAlpha.Show()
                Me.Close()

            Else

                MsgBox("Usuário ou senha inválidos!", vbExclamation, "Sistema")
                txtUsuario.Focus()
            End If
        End If
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'GeraDTI()

        txtUsuario.Text = GetCurrent.Name.Replace("AZUL\", String.Empty)
        txtDataHora.Text = Format(Now, "dd/MM/yyyy HH:mm")
        txtUsuario.CharacterCasing = CharacterCasing.Upper
    End Sub

 Private Sub frmLogin_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

 Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Login()
    End Sub

    Private Sub btnSairdoSistema_Click(sender As Object, e As EventArgs) Handles btnSairdoSistema.Click
        frmMDIAlpha.Close()
    End Sub
End Class