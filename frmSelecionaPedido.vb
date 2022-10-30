#Region "Imports"
Imports System.Data.SqlClient
Imports xAlimento
Imports GGeral
Imports System.ComponentModel
#End Region

Public Class frmSelecionaPedido

#Region "Única Instância"
    Private Shared _Instance As frmSelecionaPedido = Nothing
    Private Shared _NovaInstancia As Boolean = False

    Public Property NovaInstancia() As String
        Get
            Return _NovaInstancia
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Shared Function Instance() As frmSelecionaPedido
        If _Instance Is Nothing OrElse _Instance.IsDisposed = True Then
            _Instance = New frmSelecionaPedido
            _NovaInstancia = True
        Else
            _NovaInstancia = False
        End If
        _Instance.BringToFront()
        Return _Instance
    End Function
#End Region

#Region "Funções"

    Sub ConfiguraLista()

        'Adiciona as Colunas
        Me.lstProd.AddItemCols = 3

        'Adiciona Colunas
        Me.lstProd.Columns(0).Caption = "Código Pedido"
        Me.lstProd.Columns(1).Caption = "Código Entidade"
        Me.lstProd.Columns(2).Caption = "Fantasia"

        'Ajusta colunas
        Me.lstProd.Splits(0).DisplayColumns(0).Width = 87
        Me.lstProd.Splits(0).DisplayColumns(1).Width = 100
        Me.lstProd.Splits(0).DisplayColumns(2).Width = 300

        'Alinhamento
        Me.lstProd.Splits(0).DisplayColumns(0).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstProd.Splits(0).DisplayColumns(1).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstProd.Splits(0).DisplayColumns(2).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Near

    End Sub

    Sub AtualizaLista()
        Dim Conexao As New ConexaoSQL
        Dim drdSQL As SqlDataReader
        Dim strSQL As String
        Dim strFiltro As String = ""

        If txtCodEnt.Text = "" Then
            MsgBox("Digite um código entidade", MsgBoxStyle.Information, "ATENÇÃO")

        Else

            Try
                strSQL = " Select p.cod_vale, p.cod_ent, e.cod_ent, i.cod_vale, i.cod_prod, pd.cod_prod, pd.nome from hr.dbo.Pedido p " & vbCrLf &
                         " inner join " & vbCrLf &
                         " Entidade e On p.cod_ent = e.cod_ent " & vbCrLf &
                         " inner join " & vbCrLf &
                         " hr.dbo.Item_Ped i On p.cod_vale = i.cod_vale " & vbCrLf &
                         " inner join " & vbCrLf &
                         " Produto pd On i.cod_prod = pd.cod_prod where p.cod_vale Like '" & txtCodEnt.Text & "'"

                drdSQL = Conexao.RetornaDR(strSQL, ConexaoSQL.EnBanco.Azul)

                lstProd.ClearItems()

                Do While drdSQL.Read
                    Me.lstProd.AddItem(drdSQL("cod_vale") & ";" & drdSQL("cod_ent") & ";" & drdSQL("nome"))
                    Application.DoEvents()
                Loop

                ConfiguraLista()

                If lstProd.ListCount = 0 Then
                    lstProd.Enabled = False
                Else
                    lstProd.Enabled = True
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

#End Region

    Private Sub frmSelecionaPedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lstProd.Enabled = True

        ConfiguraLista()
    End Sub

    Private Sub frmSelecionaPedido_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        ElseIf e.KeyCode = Keys.Escape Then
            Close()
        End If
    End Sub

    Private Sub lstProd_DoubleClick(sender As Object, e As EventArgs) Handles lstProd.DoubleClick
        Dim myfrmIncluiPedido As frmIncluiPedido
        Dim Conexao As New ConexaoSQL
        Dim lngId As Long = 0

        Try
            If txtCodEnt.Text = "" Then
                MsgBox("Digite um código!", MsgBoxStyle.Critical, "ATENÇÃO!")
                Exit Sub
            End If

            myfrmIncluiPedido = frmIncluiPedido.Instance
            myfrmIncluiPedido.TipoForm = frmIncluiPedido.enumTipoForm.Altera
            myfrmIncluiPedido.RegistroAtual = lstProd.Columns(0).Value
            myfrmIncluiPedido.MdiParent = frmMDIAlpha
            myfrmIncluiPedido.Show()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnSelecionar_Click(sender As Object, e As EventArgs) Handles btnSelecionar.Click
        AtualizaLista()
    End Sub

    Private Sub btnAlterar_Click(sender As Object, e As EventArgs) Handles btnAlterar.Click
        Dim myfrmIncluiPedido As frmIncluiPedido
        Dim Conexao As New ConexaoSQL
        Dim lngId As Long = 0

        Try
            If txtCodEnt.Text = "" Then
                MsgBox("Digite o código produto!", MsgBoxStyle.Critical, "ATENÇÃO!")
                Exit Sub
            End If

            myfrmIncluiPedido = frmIncluiPedido.Instance
            myfrmIncluiPedido.TipoForm = frmIncluiPedido.enumTipoForm.Altera
            myfrmIncluiPedido.RegistroAtual = lstProd.Columns(0).Value
            myfrmIncluiPedido.MdiParent = frmMDIAlpha
            myfrmIncluiPedido.Show()

            myfrmIncluiPedido.btnGravar.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnIncluir_Click(sender As Object, e As EventArgs) Handles btnIncluir.Click
        Dim myfrmIncluiPedido As frmIncluiPedido
        Dim Conexao As New ConexaoSQL
        Dim lngId As Long = 0

        myfrmIncluiPedido = frmIncluiPedido.Instance
        myfrmIncluiPedido.TipoForm = frmIncluiPedido.enumTipoForm.Inclui
        myfrmIncluiPedido.MdiParent = frmMDIAlpha
        myfrmIncluiPedido.Show()
    End Sub

    Private Sub btnExcluir_Click(sender As Object, e As EventArgs) Handles btnExcluir.Click
        Dim Conexao As New ConexaoSQL
        Dim strSQL As String
        Dim strRegistro As String

        Try
            If lstProd.SelectedIndices.Count > 1 Then
                MsgBox("Existem mais de um registro selecionado", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub
            ElseIf Me.lstProd.SelectedIndices.Count < 1 Then
                MsgBox("Selecione um registro", MsgBoxStyle.Critical, "ERRO DE SELEÇÃO!")
                Exit Sub
            End If

            If MsgBox("Confirmar Exclusão?", MsgBoxStyle.YesNo, "EXCLUSÃO!") = MsgBoxResult.Yes Then
                lstProd.RemoveItem(0)

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnVoltar_Click(sender As Object, e As EventArgs) Handles btnVoltar.Click
        Me.Dispose()
    End Sub

    Private Sub txtCodProd_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCodEnt.KeyPress
        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
            e.Handled = True
        End If
    End Sub
End Class