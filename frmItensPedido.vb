Imports System.Data.SqlClient
Imports xAlimento
Imports GGeral
Imports System.ComponentModel

Public Class frmItensPedido

#Region "Única Instância"
    Private Shared _Instance As frmItensPedido = Nothing
    Private Shared _NovaInstancia As Boolean = False

    Public Property NovaInstancia() As String
        Get
            Return _NovaInstancia
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Shared Function Instance() As frmItensPedido
        If _Instance Is Nothing OrElse _Instance.IsDisposed = True Then
            _Instance = New frmItensPedido
            _NovaInstancia = True
        Else
            _NovaInstancia = False
        End If
        _Instance.BringToFront()
        Return _Instance
    End Function
#End Region

#Region "Declarações"
    Public Tipo As enumTipoForm
    Public Codigo_Pedido As Integer
    Enum enumTipoForm
        Incluir = 1
        Alteracao = 2
    End Enum

#End Region

#Region "Declarando Valores"
    Dim st_venda As String = "P"
    Dim cod_mestre As Integer = 0
    Dim dt_premio As Date = "1980-01-01"
    Dim caixavi As Integer = 0
    Dim caixaviVend As Integer = 0
    Dim caixates As Integer = 0
    Dim caixatesQ As Integer = 0
    Dim NotaCupom As String = "N"
    Dim retirar As Integer = 0
    Dim negociacao As Integer = 0
    Dim autorizadoatrasados As Integer = 0.00
    Dim DtImpresso As Date = "1980-01-01"
    Dim CodAutSaida As Integer = 0
    Dim NumBal As Integer = 0
    Dim Recompra As Integer = 0
    Dim PedidoCliente As String = ""
    Dim LiberadoRep As Integer = 1
    Dim DevolucaoComBoleto As Integer = 0
    Dim PgtoFaturaNumreg As Integer = 0
    Dim DtNow As Date = Now
    Dim Cod_Dav As Integer = 0
    Dim Cod_QuitaVI As Integer = 0
    Dim QuiFarelo As Integer = 0
    Dim CodChequeDev As Integer = 0
    Dim StatusMySql As String = "D"
    Dim CodEnderecoEnt As Integer = 0
    Dim DtSaida As Date = Now
    Dim LiberadoImpressao As Integer = 1
    Dim DtSaidaRep As Date = "1980-01-01"
    Dim DtAgendaCliente As Date = "1980-01-01"
    Dim ObsNota As String = "TESTE T.I"
    Dim DtPedidoSistemaRep As Date = "1980-01-01"

    Dim num_reg As Integer
    Dim cod_temp As String
    Dim avista As String = "N"
    Dim autorizado As Integer = 0.00000000
    Dim valormaximo As Integer = 0.00000000
    Dim valorminimo As Integer = 0.00000000
    Dim dtvalidade As Date = "1980-01-01"
#End Region

#Region "Funções"

#Region "Perform Click"
    Private Sub txtQuantidade_TextChanged(sender As Object, e As EventArgs) Handles txtQuantidade.TextChanged
        If Len(txtQuantidade.Text) > 0 And Len(txtValor.Text) > 0 Then
            btnMultiplica.PerformClick()
        End If
    End Sub
    Private Sub txtValor_TextChanged(sender As Object, e As EventArgs) Handles txtValor.TextChanged
        If Len(txtValor.Text) > 0 And Len(txtQuantidade.Text) > 0 Then
            btnMultiplica.PerformClick()
        End If
    End Sub

#End Region

    Sub Limpar()
        txtCodigoProduto.Text = ""
        cmbProduto.Text = ""
        txtQuantidade.Text = ""
        txtValor.Text = ""
        txtTotal.Text = ""
    End Sub

    Sub ConfiguraLista()

        'Adiciona as Colunas
        Me.lstProduto.AddItemCols = 5

        'Adiciona Colunas
        Me.lstProduto.Columns(0).Caption = "Código"
        Me.lstProduto.Columns(1).Caption = "Produto"
        Me.lstProduto.Columns(2).Caption = "Qtd"
        Me.lstProduto.Columns(3).Caption = "Valor"
        Me.lstProduto.Columns(4).Caption = "Total"

        'Ajusta colunas
        Me.lstProduto.Splits(0).DisplayColumns(0).Width = 90
        Me.lstProduto.Splits(0).DisplayColumns(1).Width = 290
        Me.lstProduto.Splits(0).DisplayColumns(2).Width = 115
        Me.lstProduto.Splits(0).DisplayColumns(3).Width = 115
        Me.lstProduto.Splits(0).DisplayColumns(4).Width = 141

        'Alinhamento
        Me.lstProduto.Splits(0).DisplayColumns(0).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstProduto.Splits(0).DisplayColumns(1).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Near
        Me.lstProduto.Splits(0).DisplayColumns(2).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstProduto.Splits(0).DisplayColumns(3).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstProduto.Splits(0).DisplayColumns(4).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
    End Sub

    Sub AtualizaLista()
        Dim Conexao As New ConexaoSQL
        Dim drdSQL As SqlDataReader
        Dim strSQL As String
        Dim strFiltro As String = ""

        Try
            If Tipo = enumTipoForm.Alteracao Then

                If cmbProduto.SelectedIndex > 0 Then _
                strFiltro &= " and cod_prod " & cmbProduto.SelectedValue

                If cmbProduto.SelectedIndex > 0 Then _
                strFiltro &= " and nome " & cmbProduto.SelectedValue

                If cmbProduto.SelectedIndex > 0 Then _
                strFiltro &= " and quantidade " & cmbProduto.SelectedValue

                If cmbProduto.SelectedIndex > 0 Then _
                strFiltro &= " and valor " & cmbProduto.SelectedValue

                If cmbProduto.SelectedIndex > 0 Then _
                strFiltro &= " total " & cmbProduto.SelectedValue

                ConfiguraLista()

                strSQL = " select i.cod_prod, i.quantidade, p.nome, p.valor from hr.dbo.Item_Ped i " & vbCrLf &
                         " inner join " & vbCrLf &
                         " Produto p on i.cod_prod = p.cod_prod where i.cod_vale = " & Codigo_Pedido

                drdSQL = Conexao.RetornaDR(strSQL, ConexaoSQL.EnBanco.Azul)

                lstProduto.ClearItems()

                Do While drdSQL.Read
                    Me.lstProduto.AddItem(drdSQL("cod_prod") & ";" & drdSQL("nome") & ";" & drdSQL("quantidade") & ";" & drdSQL("valor"))
                    Application.DoEvents()
                Loop

                ConfiguraLista()

                If lstProduto.ListCount = 0 Then
                    lstProduto.Enabled = False
                Else
                    lstProduto.Enabled = True
                End If

            Else

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub CarregaComboProduto()
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim Retorno As New DataTable

        strSQL = "select cod_prod, nome from produto where categoria like 'VENDA'"
        Retorno = Conexao.RetornaDT(strSQL, Conexao.EnBanco.Azul)

        With cmbProduto
            .DataSource = Retorno
            .DisplayMember = "Nome"
            .ValueMember = "cod_prod"
            .Splits(0).DisplayColumns(0).Visible = False
            .Splits(0).DisplayColumns(1).Width = Width - 21
            .SelectedIndex = 0
        End With
    End Sub

#End Region

#Region "Importa campos"
    Public Property cod_ent As Integer
    Public Property cod_vend As Integer
    Public Property st_pedido As String
    Public Property prazo As Integer
    Public Property dt_entrega As String
    Public Property dt_vencimento As String
    Public Property dt_lancamento As String
    Public Property dt_cobranca As String

    Public Property cod_entrega As Integer
    Public Property obs As String
    Public Property cod_recebimento As Integer
    Public Property dt_quitacao As String

#End Region

    Private Sub frmItensPedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AtualizaLista()
        ConfiguraLista()
        CarregaComboProduto()

        txtCodigoProduto.Text = ""
        cmbProduto.Text = "ESCOLHA"

        Try
            Me.Text = "Pedido - Itens"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnGravar_Click(sender As Object, e As EventArgs) Handles btnGravar.Click
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim retorno As Integer
        Dim lngId As Long = 0

        If frmItensPedido.enumTipoForm.Incluir Then

            strSQL = ("select cod_vale + 1 from Configuracao")
            retorno = Conexao.Scalar(strSQL, ConexaoSQL.EnBanco.Esta20, ConexaoSQL.EnTpRetorno.tpNumerico)
            If Conexao.Scalar(strSQL, ConexaoSQL.EnBanco.Esta20, ConexaoSQL.EnTpRetorno.tpNumerico) > 0 Then

                Try
                    If cmbProduto.SelectedValue <= 0 Then
                        MsgBox("Selecione um Produto para Efetuar o Pedido", MsgBoxStyle.Information, "ATENÇÃO!")

                    ElseIf txtQuantidade.Text = "" Then
                        MsgBox("Informe a Quantidade do Produto", MsgBoxStyle.Information, "ATENÇÃO!")

                    ElseIf txtValor.Text = "" Then
                        MsgBox("Infome o Valor do Produto", MsgBoxStyle.Information, "ATENÇÃO!")

                    ElseIf MsgBox("Deseja Efetuar o Pedido ?", MsgBoxStyle.YesNo, "ATENÇAO!") = MsgBoxResult.Yes Then

                        strSQL = String.Format(" insert into hr.dbo.Pedido (cod_vale, cod_ent, cod_vend, st_pedido, prazo, dt_entrega, dt_vencimento, dt_lancamento, dt_cobranca, cod_entrega, obs, st_venda, cod_recebimento, dt_quitacao, cod_mestre, dt_premio, caixavi, caixaviVend, caixates, caixatesQ, NotaCupom, retirar, negociacao, autorizadoatrasados, DtImpresso, CodAutSaida, NumBal, Recompra, PedidoCliente, LiberadoRep, DevolucaoComBoleto, PgtoFaturaNumreg, DtNow, Cod_Dav, Cod_QuitaVI, QuiFarelo, CodChequeDev, StatusMySql, CodEnderecoEnt, DtSaida, LiberadoImpressao, DtSaidaRep, DtAgendaCliente, ObsNota, DtPedidoSistemaRep) values " &
                                               " ({0},{1},{2},'{3}',{4},'{5}','{6}','{7}','{8}',{9},'{10}','{11}',{12},'{13}',{14},'{15}',{16},{17},{18},{19},'{20}',{21},{22},{23},'{24}',{25},{26},{27},'{28}',{29},{30},{31},'{32}',{33},{34},{35},{36},'{37}',{38},'{39}',{40},'{41}','{42}','{43}','{44}')", retorno, cod_ent, cod_vend, st_pedido, prazo, dt_entrega, dt_vencimento, dt_lancamento, dt_cobranca, cod_entrega, obs, st_venda, cod_recebimento, dt_quitacao, cod_mestre, dt_premio, caixavi, caixaviVend, caixates, caixatesQ, NotaCupom, retirar, negociacao, autorizadoatrasados, DtImpresso, CodAutSaida, NumBal, Recompra, PedidoCliente, LiberadoRep, DevolucaoComBoleto, PgtoFaturaNumreg, DtNow, Cod_Dav, Cod_QuitaVI, QuiFarelo, CodChequeDev, StatusMySql, CodEnderecoEnt, DtSaida, LiberadoImpressao, DtSaidaRep, DtAgendaCliente, ObsNota, DtPedidoSistemaRep)
                        lngId = Conexao.InsertId(strSQL, ConexaoSQL.EnBanco.Esta20)

                        If Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Esta20) Then

                            strSQL = String.Format(" insert into hr.dbo.Item_Ped (cod_vale, cod_prod, quantidade, valor, cod_temp, dtnow, avista, autorizado, valormaximo, valorminimo, dtValidade) values " & "({0},{1},{2},{3},'{4}','{5}','{6}',{7},{8},{9},'{10}')", retorno, cmbProduto.SelectedValue, txtQuantidade.Text, txtValor.Text, cod_temp, DtNow, avista, autorizado, valormaximo, valorminimo, dtvalidade)
                            lngId = Conexao.InsertId(strSQL, ConexaoSQL.EnBanco.Esta20)

                            If Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Esta20) > 0 Then
                                lstProduto.AddItem(txtCodigoProduto.Text & ";" & cmbProduto.Text & ";" & txtQuantidade.Text & ";" & txtValor.Text & ";" & txtTotal.Text)
                            End If

                            strSQL = (" update hr.dbo.Configuracao set cod_vale = cod_vale + 1")
                            Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Esta20)

                            MsgBox("Pedido Realizado com Sucesso!
                            Seu Número do Pedido é " & retorno, MsgBoxStyle.Information)
                        End If

                    Else

                        MsgBox("Erro ao Tentar efetuar o pedido", MsgBoxStyle.Critical)
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            Else
                MsgBox("Erro ao tentar efetuar o pedido !", MsgBoxStyle.Critical)
            End If

        End If
    End Sub

    Private Sub btnAlterar_Click(sender As Object, e As EventArgs) Handles btnAlterar.Click
        Dim Conexao As New ConexaoSQL
        Dim lngId As Long = 0

        Try
            If txtCodigoProduto.Text = "" And cmbProduto.SelectedValue <= 1 Then
                MsgBox("Digite um Código ou escolha Produto", MsgBoxStyle.Critical, "ATENÇÃO!")

            Else

                txtCodigoProduto.Text = lstProduto.Columns(0).Value
                cmbProduto.Text = lstProduto.Columns(1).Value
                txtQuantidade.Text = lstProduto.Columns(2).Value
                txtValor.Text = lstProduto.Columns(3).Value
                txtTotal.Text = lstProduto.Columns(4).Value
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnExlcuir_Click(sender As Object, e As EventArgs) Handles btnExlcuir.Click
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String

        Try
            If txtCodigoProduto.Text = "" And cmbProduto.SelectedValue <= 1 Then
                MsgBox("Digite um Código ou escolha Produto", MsgBoxStyle.Critical, "ATENÇÃO!")

            ElseIf MsgBox("Confirmar Exclusão?", MsgBoxStyle.YesNo, "EXCLUSÃO!") = MsgBoxResult.Yes Then
                'strSQL = ""
                'Conexao.ExecutaStr(strSQL, ConexaoSQL.EnBanco.Esta20)


                MsgBox("Excluido com sucesso", MsgBoxStyle.Information, "EXCLUSÃO!")
                lstProduto.RemoveItem(0)
                lstProduto.RemoveItem(1)
                lstProduto.RemoveItem(2)
                lstProduto.RemoveItem(3)
                txtCodigoProduto.Focus()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnVoltar_Click(sender As Object, e As EventArgs) Handles btnVoltar.Click
        Me.Dispose()
    End Sub

    Private Sub btnMultiplica_Click(sender As Object, e As EventArgs) Handles btnMultiplica.Click
        Dim MultValor As Double
        Dim MultQuantidade As Double

        MultValor = Me.txtValor.Text
        MultQuantidade = Me.txtQuantidade.Text
        Me.txtTotal.Text = MultValor * MultQuantidade
    End Sub

    Private Sub cmbProduto_TextChanged(sender As Object, e As EventArgs) Handles cmbProduto.TextChanged
        txtCodigoProduto.Text = cmbProduto.SelectedValue
    End Sub

    Private Sub lstPedido_DoubleClick(sender As Object, e As EventArgs) Handles lstProduto.DoubleClick
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String

        Try
            If lstProduto.SelectedIndices.Count > 1 Then
                MsgBox("Existem mais de um registro selecionado", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub
            ElseIf Me.lstProduto.SelectedIndices.Count < 1 Then
                MsgBox("Selecione um registro", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub

            Else

                txtCodigoProduto.Text = lstProduto.Columns(0).Value
                cmbProduto.Text = lstProduto.Columns(1).Value
                txtQuantidade.Text = lstProduto.Columns(2).Value
                txtValor.Text = lstProduto.Columns(3).Value
                txtTotal.Text = lstProduto.Columns(4).Value
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frmItensPedido_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Dispose()
        End If
    End Sub

    Private Sub txtCodigo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCodigoProduto.KeyPress
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtQuantidade_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuantidade.KeyPress
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtValor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtValor.KeyPress
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub cmbProduto_RowChange(sender As Object, e As EventArgs)
        txtTotal.Text = cmbProduto.Columns(1).Value
    End Sub

    Private Sub txtCodigo_Leave(sender As Object, e As EventArgs) Handles txtCodigoProduto.Leave
        cmbProduto.SelectedValue = txtCodigoProduto.Text
    End Sub
End Class