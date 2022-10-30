#Region "Imports"
Imports System.Data.SqlClient
Imports xAlimento
Imports GGeral
#End Region

Public Class frmAgenda

#Region "Formulário"

#Region "Única Instância"
    Private Shared _Instance As frmAgenda = Nothing
    Private Shared _NovaInstancia As Boolean = False

    Public Property NovaInstancia() As String
        Get
            Return _NovaInstancia
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Shared Function Instance() As frmAgenda
        If _Instance Is Nothing OrElse _Instance.IsDisposed = True Then
            _Instance = New frmAgenda
            _NovaInstancia = True
        Else
            _NovaInstancia = False
        End If
        _Instance.BringToFront()
        Return _Instance
    End Function
#End Region

#Region "Declarações"
    Private my_Geral As New clsGeral
    Private my_TipoForm As enumTipoForm
    Private my_CodAltera As Integer
    Public Property TipoForm() As enumTipoForm
        Get
            Return my_TipoForm
        End Get
        Set(ByVal value As enumTipoForm)
            my_TipoForm = value
        End Set
    End Property

    Enum enumTipoForm
        Altera = 1
        Inclui = 2
    End Enum

    Public Property RegistroAtual() As Integer
        Get
            Return my_CodAltera
        End Get
        Set(ByVal value As Integer)
            my_CodAltera = value
        End Set
    End Property
#End Region

    Private Sub CarregaCombo()
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim Retorno As New DataTable
        Dim strFiltro As String = ""

        strSQL = " select 0 as cod_categoria, 'ESCOLHA' as nome, 0 as ordem " & vbCrLf &
                 "union" & vbCrLf &
                 "select cod_categoria, nome, 1 as ordem from agenda_categoria order by ordem"

        Retorno = Conexao.RetornaDT(strSQL, Conexao.EnBanco.Azul)

        With cmbCategoria
            .DataSource = Retorno
            .DisplayMember = "Nome"
            .ValueMember = "cod_categoria"
            .Splits(0).DisplayColumns(0).Visible = False
            .Splits(0).DisplayColumns(1).Width = Width - 21
            .Splits(0).DisplayColumns(2).Visible = False
            .SelectedIndex = 0
        End With
    End Sub

    Private Sub frmAgenda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim drdSQL As SqlDataReader

        txtNome.CharacterCasing = CharacterCasing.Upper

        CarregaCombo()

        Try
            If Me.TipoForm = frmAgenda.enumTipoForm.Altera Then
                Me.Text = "Alteração de Agenda"
                strSQL = "select * from agenda where cod_agenda = " & Me.RegistroAtual
                drdSQL = Conexao.RetornaDR(strSQL, ConexaoSQL.EnBanco.Azul)
                If drdSQL.Read = True Then
                    lblCodigo.Value = drdSQL.Item("cod_agenda")
                    txtNome.Text = drdSQL.Item("nome")
                    txtTelefone.Text = drdSQL.Item("fone1")
                    txtEmail.Text = drdSQL.Item("email")
                    cmbCategoria.SelectedIndex = cmbCategoria.FindStringExact(drdSQL.Item("cod_categoria"), 0, 0)
                End If

            Else

                Me.Text = "Inclusão de Agenda"
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnGravar_Click(sender As Object, e As EventArgs) Handles btnGravar.Click
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim lngId As Long = 0

        Try
            If txtNome.Text = "" Then
                MsgBox("Nome não pode ser vazio", MsgBoxStyle.Critical, "ATENÇÃO!")

            ElseIf txtTelefone.Text = "" Then
                MsgBox("Telefone não pode ser vazio", MsgBoxStyle.Critical, "ATENÇÃO!")

            ElseIf txtEmail.Text = "" Then
                MsgBox("Email não pode ser vazio", MsgBoxStyle.Critical, "ATENÇÃO!")

            ElseIf cmbCategoria.SelectedIndex <= 0 Then
                MsgBox("Selecione uma Categoria", MsgBoxStyle.Critical, "ATENÇÃO!")

            Else

                If Me.TipoForm = frmAgenda.enumTipoForm.Inclui Then
                    If MsgBox("Adicionar contato ?", MsgBoxStyle.YesNo, "INCLUSÃO") = MsgBoxResult.Yes Then
                        strSQL = String.Format(" insert into agenda (nome, fone1, email, cod_categoria) values " &
                                                               " ('{0}','{1}','{2}',{3})", Me.txtNome.Text, Me.txtTelefone.Text, Me.txtEmail.Text, cmbCategoria.SelectedValue)
                        lngId = Conexao.InsertId(strSQL, ConexaoSQL.EnBanco.Azul)
                        MsgBox("Inclusão de cadastro concluída: " & lngId, MsgBoxStyle.Information)
                        txtNome.Focus()
                    End If

                Else

                    If MsgBox("Confirmar Alteração?", MsgBoxStyle.YesNo, "ALTERAÇÃO") = MsgBoxResult.Yes Then
                        strSQL = "update agenda set nome = '" & Me.txtNome.Text & "',fone1 = '" & Me.txtTelefone.Text & "', email = '" & Me.txtEmail.Text & "' where cod_agenda = " & Me.lblCodigo.Text
                        Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Azul)
                        lngId = Conexao.InsertId(strSQL, ConexaoSQL.EnBanco.Azul)
                        MsgBox("Alteração Concluída!")
                    End If

                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnVoltar_Click(sender As Object, e As EventArgs) Handles btnVoltar.Click
        Me.Dispose()
        Exit Sub
    End Sub


    Private Sub frmAgenda_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If

        If e.KeyCode = Keys.Escape Then
            Me.Dispose()
        End If
    End Sub

    Private Sub txtNome_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNome.KeyPress
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Then
            e.Handled = True
            e = Nothing
        End If
    End Sub

    Private Sub txtTelefone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTelefone.KeyPress
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtEmail_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEmail.KeyPress
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Then
            e.Handled = True
            e = Nothing
        End If
    End Sub
#End Region
End Class
