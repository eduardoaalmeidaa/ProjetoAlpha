#Region "Imports"
Imports xAlimento
Imports GGeral
Imports System.ComponentModel
#End Region

#Region "Imports Relatorio"
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing.Printing
#End Region

Public Class frmRelatorio

#Region "FUNÇÃO-LISTA"
    Sub ConfiguraLista()

        'Adiciona as Colunas
        Me.lstAgenda.AddItemCols = 4

        'Adiciona Colunas
        Me.lstAgenda.Columns(0).Caption = "Código"
        Me.lstAgenda.Columns(1).Caption = "Nome"
        Me.lstAgenda.Columns(2).Caption = "Telefone"
        Me.lstAgenda.Columns(3).Caption = "Email"

        'Ajusta colunas
        Me.lstAgenda.Splits(0).DisplayColumns(0).Width = 50
        Me.lstAgenda.Splits(0).DisplayColumns(1).Width = 200
        Me.lstAgenda.Splits(0).DisplayColumns(2).Width = 120
        Me.lstAgenda.Splits(0).DisplayColumns(3).Width = 200

        'Alinhamento
        Me.lstAgenda.Splits(0).DisplayColumns(0).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstAgenda.Splits(0).DisplayColumns(1).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Near
        Me.lstAgenda.Splits(0).DisplayColumns(2).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
        Me.lstAgenda.Splits(0).DisplayColumns(3).Style.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Near
    End Sub

    Sub Atualizalista()
        Dim Conexao As New ConexaoSQL
        Dim drdSQL As SqlDataReader
        Dim strSQL As String
        Dim strFiltro As String = ""
        Try
            ConfiguraLista()

            strSQL = "select * from agenda where cod_agenda > 0" & strFiltro

            drdSQL = Conexao.RetornaDR(strSQL, ConexaoSQL.EnBanco.Azul)

            lstAgenda.ClearItems()

            Do While drdSQL.Read
                Me.lstAgenda.AddItem(drdSQL("cod_agenda") & ";" & drdSQL("nome") & ";" & drdSQL("fone1") & ";" & drdSQL("email"))
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

#End Region

    'Definição das variáveis visiveis no formulário
    Private conexaoNorthwind As SqlConnection
    Private Leitor As SqlDataReader
    Dim cmd As SqlCommand
    Private paginaAtual As Integer = 1
    Private RelatorioTítulo As String

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        AgendaSelecao.Imprimir("")
    End Sub

    Private Sub btnVoltar_Click(sender As Object, e As EventArgs) Handles btnVoltar.Click
        Me.Dispose()
    End Sub

    Private Sub frmRelatorio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfiguraLista()
        Atualizalista()
    End Sub

    Private Sub frmRelatorio_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Dispose()
        End If
    End Sub

End Class