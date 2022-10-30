#Region "Imports"
Imports System.Data.SqlClient
Imports xAlimento
Imports GGeral
Imports System.ComponentModel
#End Region

Public Class frmIncluiPedido

#Region "Única Instância"
    Private Shared _Instance As frmIncluiPedido = Nothing
    Private Shared _NovaInstancia As Boolean = False

    Public Property NovaInstancia() As String
        Get
            Return _NovaInstancia
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Shared Function Instance() As frmIncluiPedido
        If _Instance Is Nothing OrElse _Instance.IsDisposed = True Then
            _Instance = New frmIncluiPedido
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
    Private GetCurrent As Object

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

#Region "Funções"

    Sub CarregaComboRepresentante()
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim Retorno As New DataTable

        strSQL = "select * from Vendedor"
        Retorno = Conexao.RetornaDT(strSQL, Conexao.EnBanco.Azul)

        With cmbRepresentante
            .DataSource = Retorno
            .DisplayMember = "Nome"
            .ValueMember = "cod_vend"
            .Splits(0).DisplayColumns(0).Visible = False
            .Splits(0).DisplayColumns(1).Width = Width - 21
            .Splits(0).DisplayColumns(2).Visible = False
            .SelectedIndex = 0
        End With
    End Sub

    Sub CarregaComboRecebimento()
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim Retorno As New DataTable

        strSQL = "select * from Recebimento where status_palm like '%s%'"
        Retorno = Conexao.RetornaDT(strSQL, Conexao.EnBanco.Azul)

        With cmbRecebimento
            .DataSource = Retorno
            .DisplayMember = "Nome"
            .ValueMember = "cod_recebimento"
            .Splits(0).DisplayColumns(0).Visible = False
            .Splits(0).DisplayColumns(1).Width = Width - 21
            .Splits(0).DisplayColumns(2).Visible = False
            .SelectedIndex = 0
        End With
    End Sub

#End Region

    Private Sub frmIncluiPedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim drdSQL As SqlDataReader

        With cmbStatusPedido
            .AddItem("ESCOLHA")
            .AddItem("V- CONFERIDO;CONFERIDO")
            .AddItem("D- DISPONÍVEL;DISPONÍVEL")
            .AddItem("Q- QUITADO;QUITADO")
            .AddItem("C- CANCELADO;CANCELADO")
            .AddItem("T- TROCA CONFERIDA;TROCA CONFERIDA")
            .AddItem("B- BONIFICAÇÃO CONFERIDA;BONIFICAÇÃO CONFERIDA")
            .AddItem("X- NÃO IMPRESSO;NÃO IMPRESSO")
            .AddItem("I- IMPRESSO;IMPRESSO")
            .DisplayMember = "status"
            .ValueMember = "status"
            .Splits(0).DisplayColumns(0).Visible = True
            .Splits(0).DisplayColumns(0).Width = Width - 21
            .Splits(0).DisplayColumns(1).Visible = False
            .SelectedIndex = 0
        End With

        CarregaComboRepresentante()

        cmbStatusPedido.Text = "ESCOLHA"

        cmbRepresentante.Text = "ESCOLHA"

        CarregaComboRecebimento()

        Try
            If Me.TipoForm = frmIncluiPedido.enumTipoForm.Altera Then
                strSQL = " select p.cod_vale, p.cod_ent, p.cod_vend, p.st_pedido, p.prazo, p.dt_entrega, p.dt_vencimento, p.cod_entrega,
                           p.st_recebimento, p.dt_quitacao, i.cod_vale, i.cod_prod, i.quantidade, i.valor from hr.dbo.Pedido p " & vbCrLf &
                         " inner join " & vbCrLf &
                         " hr.dbo.Item_Ped i on p.cod_vale = i.cod_vale where p.cod_vale = " & Me.RegistroAtual
                drdSQL = Conexao.RetornaDR(strSQL, ConexaoSQL.EnBanco.Azul)
                If drdSQL.Read = True Then
                    lblNPedido.Text = drdSQL.Item("cod_vale")
                    txtCodigo.Text = drdSQL.Item("cod_ent")
                    cmbRepresentante.SelectedIndex = cmbRepresentante.FindStringExact(drdSQL.Item("cod_vend"), 0, 0)
                    cmbStatusPedido.SelectedIndex = cmbStatusPedido.FindStringExact(drdSQL.Item("st_pedido"), 0, 0)
                    txtPrazo.Text = drdSQL.Item("prazo")
                    DateEditEntrega.Value = drdSQL.Item("dt_entrega")
                    DateEditVencimento.Value = drdSQL.Item("dt_vencimento")
                    DateEditLancamento.Value = drdSQL.Item("dt_lancamento")
                    txtNEntrega.Text = drdSQL.Item("cod_entrega")
                    cmbRecebimento.SelectedIndex = cmbRecebimento.FindStringExact(drdSQL.Item("st_recebimento"))
                    DateEditQuitacao.Value = drdSQL.Item("dt_quitacao")
                End If
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
            If txtCodigo.Text = "" Then
                MsgBox("Código não pode ser vazio", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf DateEditEntrega.Text = "" Then
                MsgBox("Selecione a data entrega!", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf DateEditVencimento.Text = "" Then
                MsgBox("Selecione a data de vencimento!", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf DateEditLancamento.Text = "" Then
                MsgBox("Selecione a data de lançamento!", MsgBoxStyle.Information, "ATENÇÃO!")

                'ElseIf DateEditQuitacao.Text = "" Then
                '    MsgBox("Selecione a data de quitação!", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf cmbRepresentante.SelectedIndex < 0 Then
                MsgBox("Selecione o representante!", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf cmbRecebimento.SelectedIndex <= 0 Then
                MsgBox("Selecione o método de recebimento", MsgBoxStyle.Information, "ATENÇÃO!")

                'Else

                '    If MsgBox("Gravar pedido?", MsgBoxStyle.YesNo, "INCLUSÃO DE PEDIDO") = MsgBoxResult.Yes Then
                '        strSQL = String.Format(" insert into hr.dbo.Pedido (cod_vale, cod_ent, cod_vend, st_pedido, prazo, dt_entrega, dt_vencimento, dt_lancamento, obs, cod_recebimento, dt_quitacao) values " &
                '                               " ({0},{1},{2},{3},'{4}','{5}',{6},'{7}',{8},{9})", Me.lblNPedido.Text, Me.txtCodigo.Text, Me.cmbRepresentante.SelectedValue,
                '                                                                                   Me.cmbStatusPedido.Text, Me.txtPrazo.Text, Me.DateEditEntrega.Text,
                '                                                                                   Me.DateEditVencimento.Text, Me.DateEditLancamento.Text, Me.txtObservacaodoPedido.Text,
                '                                                                                   Me.cmbRecebimento.SelectedValue, Me.DateEditQuitacao.Text)

                '        lngId = Conexao.InsertId(strSQL, ConexaoSQL.EnBanco.Azul)
                '        MsgBox("Inclusão de pedido concluída: " & lngId, MsgBoxStyle.Information)
                '    End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnItens_Click(sender As Object, e As EventArgs) Handles btnItens.Click
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim drdSQL As SqlDataReader
        Dim lngId As Long = 0

        Try
            If txtCodigo.Text = "" Then
                MsgBox("Código não pode ser vazio", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf DateEditEntrega.Text = "" Then
                MsgBox("Selecione a data entrega!", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf DateEditVencimento.Text = "" Then
                MsgBox("Selecione a data de vencimento!", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf DateEditLancamento.Text = "" Then
                MsgBox("Selecione a data de lançamento!", MsgBoxStyle.Information, "ATENÇÃO!")

                'ElseIf DateEditQuitacao.Text = "" Then
                '    MsgBox("Selecione a data de quitação!", MsgBoxStyle.Information, "ATENÇÃO!")

            ElseIf cmbRepresentante.SelectedIndex < 0 Then
                MsgBox("Selecione o representante!", MsgBoxStyle.Information, "ATENÇÃO!")


            ElseIf cmbRecebimento.SelectedIndex <= 0 Then
                MsgBox("Selecione o método de recebimento", MsgBoxStyle.Information, "ATENÇÃO!")

            Else

                Dim frmItensPedido As frmItensPedido
                frmItensPedido = frmItensPedido.Instance
                frmItensPedido.cod_ent = txtCodigo.Text
                frmItensPedido.cod_vend = cmbRepresentante.SelectedValue
                frmItensPedido.st_pedido = cmbStatusPedido.SelectedText(0)
                frmItensPedido.prazo = txtPrazo.Text
                frmItensPedido.dt_entrega = DateEditEntrega.SelectedText
                frmItensPedido.dt_vencimento = DateEditVencimento.SelectedText
                frmItensPedido.dt_lancamento = DateEditLancamento.SelectedText
                frmItensPedido.obs = txtObservacaodoPedido.Text
                frmItensPedido.cod_recebimento = cmbRecebimento.SelectedValue
                frmItensPedido.dt_quitacao = DateEditQuitacao.SelectedText
                frmItensPedido.dt_cobranca = DateEditCobranca.SelectedText
                frmItensPedido.cod_entrega = txtNEntrega.Text
                frmItensPedido.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnVoltar_Click(sender As Object, e As EventArgs) Handles btnVoltar.Click
        Me.Dispose()
    End Sub

    Private Sub txtCodigo_TextChanged(sender As Object, e As EventArgs) Handles txtCodigo.TextChanged
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim drdSQL As SqlDataReader

        If txtCodigo.Text <> "" Then
            Try
                strSQL = " select p.cod_vale, p.cod_ent, p.st_pedido, p.prazo, p.dt_entrega, p.dt_vencimento, p.dt_lancamento,
                           p.cod_entrega, p.obs, p.cod_recebimento, p.dt_quitacao, p.cod_coneccao, e.cod_ent, e.fantasia, e.razao, e.endereco,
                           e.cod_bairro, e.obs, e.cod_recebimento, e.prazo, b.nome, b.cod_cidade, c.CODIGO, c.CIDADE from hr.dbo.pedido p " & vbCrLf &
                         " inner join " & vbCrLf &
                         " Entidade e on p.cod_ent = e.cod_ent " & vbCrLf &
                         " inner join " & vbCrLf &
                         " Bairro b on e.cod_bairro = b.cod_bairro " & vbCrLf &
                         " inner join " & vbCrLf &
                         " Cidade c on b.cod_cidade = c.CODIGO where e.cod_ent = " & txtCodigo.Text
                drdSQL = Conexao.RetornaDR(strSQL, ConexaoSQL.EnBanco.Azul)

                If drdSQL.Read = True Then
                    txtCodigo.Text = drdSQL.Item("cod_ent")
                    lblFantasia.Text = drdSQL.Item("fantasia")
                    lblRazao.Text = drdSQL.Item("razao")
                    lblCidade.Text = drdSQL.Item("CIDADE")
                    lblEndereco.Text = drdSQL.Item("endereco")
                    lblBairro.Text = drdSQL.Item("nome")
                    lblObsEntidade.Text = drdSQL.Item("obs")
                    DateEditLancamento.Text = Now
                    txtPrazo.Text = drdSQL.Item("prazo")
                    cmbRecebimento.SelectedIndex = drdSQL.Item("cod_recebimento")
                    cmbStatusPedido.SelectedIndex = 2
                    txtNEntrega.Text = 999999

                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub frmIncluiPedido_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Dispose()
        End If
    End Sub

    Private Sub txtCodigo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCodigo.KeyPress
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPrazo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrazo.KeyPress
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtDataEntrega_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNEntrega_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNEntrega.KeyPress
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtSomaData_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSomaData.KeyPress
        Dim Entrega As Date = Date.Parse(DateEditEntrega.Value)
        Dim Prazo As Integer = Integer.Parse(txtPrazo.Text)

        DateEditVencimento.Value = DateAdd(DateInterval.Day, Prazo, Entrega)

        'Try
        '    If txtCodigo.Text = "" Then
        '        MsgBox("Digite um Código!1", MsgBoxStyle.Information, "ATENÇÃO!")

        '    ElseIf txtPrazo.Text = "" Then
        '        MsgBox("Digite um Código!2", MsgBoxStyle.Information, "ATENÇÃO!")

        '    ElseIf DateEditEntrega.Text <> "" Then
        '        MsgBox("Digite um Código!3", MsgBoxStyle.Information, "ATENÇÃO!")
        '    End If

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub
End Class