#Region "Imports"
Imports System.Data.SqlClient
Imports xAlimento
Imports GGeral
Imports System.ComponentModel
#End Region

Public Class frmSeleciona

#Region "Formulario"

#Region "Única Instância"
    Private Shared _Instance As frmSeleciona = Nothing
    Private Shared _NovaInstancia As Boolean = False

    Public Property NovaInstancia() As String
        Get
            Return _NovaInstancia
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Shared Function Instance() As frmSeleciona
        If _Instance Is Nothing OrElse _Instance.IsDisposed = True Then
            _Instance = New frmSeleciona
            _NovaInstancia = True
        Else
            _NovaInstancia = False
        End If
        _Instance.BringToFront()
        Return _Instance
    End Function
#End Region

#Region "TIMER ListaAgenda"
    Private Sub TimerAtualizaLista_Tick(sender As Object, e As EventArgs) Handles TimerAtualizaLista.Tick
        Atualizalista()
    End Sub

    Private Sub TimerCarregaCombo_Tick(sender As Object, e As EventArgs) Handles TimerCarregaCombo.Tick
        CarregaCombo()
    End Sub
#End Region

#Region "Funções"
    Sub ConfiguraLista()

        'Adiciona as Colunas
        Me.lstAgenda.AddItemCols = 5

        'Adiciona Colunas
        Me.lstAgenda.Columns(0).Caption = "Código"
        Me.lstAgenda.Columns(1).Caption = "Nome"
        Me.lstAgenda.Columns(2).Caption = "Telefone"
        Me.lstAgenda.Columns(3).Caption = "Email"
        Me.lstAgenda.Columns(4).Caption = "Categoria"

        'Ajusta colunas
        Me.lstAgenda.Splits(0).DisplayColumns(0).Width = 50
        Me.lstAgenda.Splits(0).DisplayColumns(1).Width = 200
        Me.lstAgenda.Splits(0).DisplayColumns(2).Width = 120
        Me.lstAgenda.Splits(0).DisplayColumns(3).Width = 200
        Me.lstAgenda.Splits(0).DisplayColumns(4).Width = 200

        'Alinhamento
        Me.lstAgenda.Splits(0).DisplayColumns(0).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstAgenda.Splits(0).DisplayColumns(1).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Near
        Me.lstAgenda.Splits(0).DisplayColumns(2).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstAgenda.Splits(0).DisplayColumns(3).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Near
        Me.lstAgenda.Splits(0).DisplayColumns(4).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center

    End Sub

    Sub Atualizalista()
        Dim Conexao As New ConexaoSQL
        Dim drdSQL As SqlDataReader
        Dim strSQL As String
        Dim strFiltro As String = ""

        Try
            If txtNome.Text <> "" Then _
                strFiltro &= " and nome like '" & txtNome.Text & "%' "

            If txtTelefone.Text <> "" Then _
                strFiltro &= " and fone1 like '" & txtTelefone.Text & "%' "

            If txtEmail.Text <> "" Then _
                 strFiltro &= " and email like '" & txtEmail.Text & "%' "

            If cmbCategoria.SelectedIndex > 0 Then _
                strFiltro &= " and cod_categoria = " & cmbCategoria.SelectedValue

            ConfiguraLista()

            strSQL = "select * from agenda where cod_agenda > 0" & strFiltro

            drdSQL = Conexao.RetornaDR(strSQL, ConexaoSQL.EnBanco.Azul)

            lstAgenda.ClearItems()

            Do While drdSQL.Read
                Me.lstAgenda.AddItem(drdSQL("cod_agenda") & ";" & drdSQL("nome") & ";" & drdSQL("fone1") & ";" & drdSQL("email") & ";" & drdSQL("cod_categoria"))
                Application.DoEvents()
            Loop

            ConfiguraLista()

            If lstAgenda.ListCount = 0 Then
                lstAgenda.Enabled = False
            Else
                lstAgenda.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub CarregaCombo()
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim Retorno As New DataTable

        strSQL = "select 0 as cod_categoria, 'ESCOLHA' as nome, 0 as ordem" & vbCrLf &
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

#Region "Perform click"

    Private Sub txtNome_TextChanged(sender As Object, e As EventArgs) Handles txtNome.TextChanged
        If Len(txtNome.Text) > 0 Then
            btnSelect.PerformClick()
        End If
    End Sub

    Private Sub txtTelefone_TextChanged(sender As Object, e As EventArgs) Handles txtTelefone.TextChanged
        If Len(txtTelefone.Text) > 0 Then
            btnSelect.PerformClick()
        End If
    End Sub

    Private Sub txtEmail_TextChanged(sender As Object, e As EventArgs) Handles txtEmail.TextChanged
        If Len(txtEmail.Text) > 0 Then
            btnSelect.PerformClick()
        End If
    End Sub
#End Region

#End Region

    Private Sub frmSeleciona_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TimerAtualizaLista.Enabled = True
        TimerCarregaCombo.Enabled = True

        CarregaCombo()
        ConfiguraLista()

        lstAgenda.Enabled = True

        txtNome.CharacterCasing = CharacterCasing.Upper
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        Atualizalista()
    End Sub

    Private Sub btnAlterar_Click(sender As Object, e As EventArgs) Handles btnAlterar.Click
        Dim myfrmAgenda As frmAgenda
        Dim Conexao As New ConexaoSQL
        Dim lngId As Long = 0

        Try
            If lstAgenda.SelectedIndices.Count > 1 Then
                MsgBox("Existem mais de um registro selecionado", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub

            ElseIf Me.lstAgenda.SelectedIndices.Count < 1 Then
                MsgBox("Selecione um registro", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub
            End If

            myfrmAgenda = frmAgenda.Instance
            myfrmAgenda.TipoForm = frmAgenda.enumTipoForm.Altera
            myfrmAgenda.RegistroAtual = lstAgenda.Columns(0).Value
            myfrmAgenda.MdiParent = frmMDIAlpha
            myfrmAgenda.Show()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnExcluir_Click(sender As Object, e As EventArgs) Handles btnExcluir.Click
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim strRegistro As String

        Try
            If lstAgenda.SelectedIndices.Count > 1 Then
                MsgBox("Existem mais de um registro selecionado", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub
            ElseIf Me.lstAgenda.SelectedIndices.Count < 1 Then
                MsgBox("Selecione um registro", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub
            End If

            If MsgBox("Confirmar Exclusão?", MsgBoxStyle.YesNo, "EXCLUSÃO!") = MsgBoxResult.Yes Then
                strRegistro = lstAgenda.Columns(0).Value
                strSQL = " delete agenda where cod_agenda = " & strRegistro
                Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Azul)

                Atualizalista()

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnLimpar_Click(sender As Object, e As EventArgs) Handles btnLimpar.Click
        txtNome.Text = ""
        txtTelefone.Text = ""
        txtEmail.Text = ""
        txtNome.Focus()
        CarregaCombo()
    End Sub

    Private Sub btnVoltar_Click(sender As Object, e As EventArgs) Handles btnVoltar.Click
        Me.Dispose()
    End Sub

    Private Sub lstAgenda_DoubleClick(sender As Object, e As EventArgs) Handles lstAgenda.DoubleClick
        Dim myfrmAgenda As frmAgenda
        Dim Conexao As New ConexaoSQL
        Dim lngId As Long = 0

        Try
            If lstAgenda.SelectedIndices.Count > 1 Then
                MsgBox("Existem mais de um registro selecionado", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub
            ElseIf Me.lstAgenda.SelectedIndices.Count < 1 Then
                MsgBox("Selecione um registro", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub
            End If

            myfrmAgenda = frmAgenda.Instance
            myfrmAgenda.TipoForm = frmAgenda.enumTipoForm.Altera
            myfrmAgenda.RegistroAtual = lstAgenda.Columns(0).Value
            myfrmAgenda.MdiParent = frmMDIAlpha
            myfrmAgenda.Show()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frmSeleciona_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Dispose()
        End If
    End Sub

    Private Sub txtNome_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNome.KeyPress
        'Permite somente a entrada de Caracteres
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

