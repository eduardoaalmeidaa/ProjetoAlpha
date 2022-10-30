#Region "Imports"
Imports System.Data.SqlClient
Imports xAlimento
Imports GGeral
#End Region

Public Class frmCategoria

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

        strSQL = " select 0 as cod_categoria, '' as nome, 0 as ordem " & vbCrLf &
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

    Private Sub frmCategoria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtCategoria.CharacterCasing = CharacterCasing.Upper

        If txtCategoria.Text = "" Then
            cmbCategoria.SelectedValue = 0
        End If

        CarregaCombo()

        cmbCategoria.SelectedText = "ESCOLHA"
    End Sub

    Private Sub btnIncluir_Click(sender As Object, e As EventArgs) Handles btnIncluir.Click
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String

        If txtCategoria.Text <> "" And txtCategoria.Text <> "ESCOLHA" Then

            strSQL = "select Count(*) from Agenda_Categoria where nome like '" & txtCategoria.Text & "%'"
            If Conexao.Scalar(strSQL, ConexaoSQL.EnBanco.Azul, ConexaoSQL.EnTpRetorno.tpNumerico) > 0 Then
                xMsgBox.Show("Categoria já cadastrada", xMsgBox.Contexto.Warning)

            ElseIf MsgBox("Adicionar Categoria ?", MsgBoxStyle.YesNo, "INCLUSÃO DE CATEGORIA") = MsgBoxResult.Yes Then
                strSQL = " insert into agenda_categoria (nome) values ('" & txtCategoria.Text & "')"
                Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Azul)
                MsgBox("Categoria adicionada com sucesso!", MsgBoxStyle.Information, "CATEGORIA")
                CarregaCombo()
                txtCategoria.Focus()
            End If

        Else

            MsgBox("Categoria não pode ser vazio!", MsgBoxStyle.Critical, "CATEGORIA")
        End If
    End Sub

    Private Sub btnAlterar_Click(sender As Object, e As EventArgs) Handles btnAlterar.Click
        Dim Conexao As New ConexaoSQL
        Dim lngId As Long = 0
        Dim strSQL As String

        If txtCategoria.Text <> "" And txtCategoria.Text <> "ESCOLHA" Then

            strSQL = "select Count(*) from Agenda_Categoria where nome like '" & txtCategoria.Text & "%'"
            If Conexao.Scalar(strSQL, ConexaoSQL.EnBanco.Azul, ConexaoSQL.EnTpRetorno.tpNumerico) > 0 Then
                xMsgBox.Show("Ambos campos não podem ser iguais!", xMsgBox.Contexto.Warning)

            ElseIf MsgBox("Alterar Categoria?", MsgBoxStyle.YesNo, "ALTERAÇÃO DE CATEGORIA") = MsgBoxResult.Yes Then
                strSQL = " update Agenda_Categoria set nome = '" & txtCategoria.Text & "' where cod_categoria = '" & cmbCategoria.SelectedValue & "' "
                Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Azul)
                MsgBox("Categoria alterada com sucesso!", MsgBoxStyle.Information, "CATEGORIA")

                CarregaCombo()
                txtCategoria.Focus()
            End If

        Else

            MsgBox("Categoria não pode ser vazio!", MsgBoxStyle.Critical, "CATEGORIA")

        End If
    End Sub

    Private Sub btnExcluir_Click(sender As Object, e As EventArgs) Handles btnExcluir.Click
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim strCategoria As String

        Try
            If Me.cmbCategoria.SelectedValue < 1 Then
                MsgBox("Selecione uma Categoria para Excluir", MsgBoxStyle.Critical, "ERRO")

            Else

                strSQL = " select Count(*) from Agenda where cod_categoria like '" & cmbCategoria.SelectedValue & "%'"
                If Conexao.Scalar(strSQL, ConexaoSQL.EnBanco.Azul, ConexaoSQL.EnTpRetorno.tpNumerico) > 0 Then
                    xMsgBox.Show("Existem Contatos com essa Categoria !", xMsgBox.Contexto.Warning)

                ElseIf MsgBox("Deseja Excluir Categoria ?", MsgBoxStyle.YesNo, "Exclusão") = MsgBoxResult.Yes Then
                    strCategoria = cmbCategoria.SelectedValue
                    strSQL = " delete agenda_categoria where cod_categoria = " & strCategoria
                    Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Azul)
                    xMsgBox.Show("Categoria Excluída com Sucesso !", xMsgBox.Contexto.Warning)
                    CarregaCombo()
                    txtCategoria.Focus()
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnVoltar_Click(sender As Object, e As EventArgs) Handles btnVoltar.Click
        Me.Dispose()
    End Sub

    Private Sub frmCategoria_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnIncluir.Focus()

            If txtCategoria.Text = "" Then
                MsgBox("Categoria não pode ser vazio!", MsgBoxStyle.Critical, "CATEGORIA")
            End If

        End If

        If e.KeyCode = Keys.Enter Then
            SendKeys.Send(" {TAB} ")
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Dispose()
        End If
    End Sub

    Private Sub cmbCategoria_RowChange(sender As Object, e As EventArgs) Handles cmbCategoria.RowChange
        txtCategoria.Text = cmbCategoria.Columns(1).Value
    End Sub
#End Region
End Class