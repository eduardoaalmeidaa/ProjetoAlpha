#Region "Imports"
Imports Negocios
Imports TratamentoDeErros
Imports GGeral
Imports xAlimento
Imports System.Drawing.Printing
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Reflection
#End Region
Public Class AgendaSelecao

#Region "Declarações"
    Private Shared Instancia As AgendaSelecao
    Private ReadOnly CfgRel As New CfgRelatorio(CfgRelatorio.EnOrientacao.RETRATO)
    Private WithEvents Documento As New PrintDocument
    Private SetTamanhoPapel As Boolean = False
    Private AgendaSelecao As New DataTable
    Private Ponteiro As Integer = 0

    Private ColunaTituloRel As Integer
    Private ColunaCodigo As Integer
    Private ColunaNome As Integer
    Private ColunaTelefone As Integer
    Private ColunaEmail As Integer
    Private ColunaDataImpressao As Integer
#End Region

#Region "Métodos"
    Public Shared Sub Imprimir(ByRef Filtro As String)
1:      Instancia = New AgendaSelecao()
6:      clsGeral.ImprimirDocumento(Instancia.Documento, "Agenda Seleção", False, False)
    End Sub

    Private Sub CfgVariaveisLayout()
1:      Try
2:          CfgRel.MargemEsquerda = 4
3:          CfgRel.MargemDireita = 206
4:          CfgRel.MargemTopo = 8
5:          CfgRel.MargemRodape = 278

6:          CfgRel.PosicaoLinha = CfgRel.MargemTopo

            ColunaTituloRel = CfgRel.MargemEsquerda
            ColunaCodigo = CfgRel.MargemEsquerda
            ColunaNome = CfgRel.MargemEsquerda + 10
            ColunaTelefone = CfgRel.MargemEsquerda + 85
            ColunaEmail = CfgRel.MargemEsquerda + 130
            ColunaDataImpressao = CfgRel.MargemEsquerda + 30

            CfgRel.NumeroPagina = 1

7:      Catch ex As Exception
8:          Erro.Log(ConexaoSQL.VersaoAplicacao, Assembly.GetExecutingAssembly.GetName.Name(), ex.Message &
                     " " & ex.StackTrace, ToString, MethodBase.GetCurrentMethod().Name, True, Erl.ToString)
9:      End Try
    End Sub
    Private Sub PreencheDados()
1:      Dim StrSQL As String = ""
2:      Dim Conexao As New ConexaoSQL
3:      Dim Retorno As New DataTable
4:      Try
5:          StrSQL = "select cod_agenda as Codigo, Nome, fone1 as Telefone, Email from Agenda "

6:          Retorno = Conexao.RetornaDT(StrSQL, ConexaoSQL.EnBanco.Azul)

7:          If IsNothing(Retorno) OrElse Retorno.Rows.Count = 0 Then
8:              xMsgBox.Show("Não foi possivel retornar a consulta do banco, entre em contato com T.I(Tecnologia da Informação)", xMsgBox.Contexto.Erro)
9:              Exit Sub
10:         End If

11:         AgendaSelecao = Retorno.Copy
12:     Catch ex As Exception
13:         Erro.Log(ConexaoSQL.VersaoAplicacao, Assembly.GetExecutingAssembly.GetName.Name(), ex.Message &
                     " " & ex.StackTrace, ToString, MethodBase.GetCurrentMethod().Name, True, Erl.ToString)
14:     End Try
    End Sub
#End Region

#Region "Eventos Impressao"
    Private Sub Documento_BeginPrint(sender As Object, e As PrintEventArgs) Handles Documento.BeginPrint
        Try
            CfgVariaveisLayout()
        Catch ex As Exception
            Erro.Log(ConexaoSQL.VersaoAplicacao, Assembly.GetExecutingAssembly.GetName.Name(), ex.Message &
                     " " & ex.StackTrace, ToString, MethodBase.GetCurrentMethod().Name, True, Erl.ToString)
        End Try
    End Sub

    Private Sub Documento_PrintPage(sender As Object, e As PrintPageEventArgs) Handles Documento.PrintPage
        Try
            With e
                .Graphics.PageUnit = GraphicsUnit.Millimeter
                If SetTamanhoPapel = False Then
                    CfgRelatorio.AlteraTamanhoPapel(e, PaperKind.A4)
                    SetTamanhoPapel = True
                End If

                If IsNothing(AgendaSelecao) OrElse AgendaSelecao.Rows.Count = 0 Then
                    PreencheDados()
                    If IsNothing(AgendaSelecao) OrElse AgendaSelecao.Rows.Count = 0 Then
                        e.HasMorePages = False
                        Exit Sub
                    End If
                End If

                ImprimeCabecalho(e)
                ImprimeCorpo(e)
            End With
        Catch ex As Exception
            Erro.Log(ConexaoSQL.VersaoAplicacao, Assembly.GetExecutingAssembly.GetName.Name(), ex.Message &
                     " " & ex.StackTrace, ToString, MethodBase.GetCurrentMethod().Name, True, Erl.ToString)
        End Try
    End Sub
    Private Sub Documento_EndPrint(sender As Object, e As PrintEventArgs) Handles Documento.EndPrint
        AgendaSelecao = Nothing
        CfgRel.NumeroPagina = 1
        Ponteiro = 0
    End Sub
#End Region

#Region "Metodos Impressao"
    Private Sub ImprimeCabecalho(e As PrintPageEventArgs)
        Try
            With e.Graphics
                CfgRel.PosicaoLinha += 1
                .DrawString("Relatório de Agenda", CfgRelatorio.Fontes.Arial8B, Brushes.Black, ColunaTituloRel, CfgRel.PosicaoLinha)
                .DrawString("Pag. " & CfgRel.NumeroPagina, CfgRelatorio.Fontes.Arial7, Brushes.Black, CfgRel.MargemEsquerda + 187, CfgRel.PosicaoLinha)

                CfgRel.PosicaoLinha += 4
                CfgRelatorio.LinhaHorizontal(CfgRelatorio.EnTipoLinha.Fina, CfgRel.MargemEsquerda, CfgRel.MargemDireita, CfgRel.PosicaoLinha, e)
                CfgRel.PosicaoLinha += 2

                .DrawString("Código", CfgRelatorio.Fontes.Arial7B, Brushes.Black, ColunaCodigo, CfgRel.PosicaoLinha)
                .DrawString("Nome", CfgRelatorio.Fontes.Arial7B, Brushes.Black, ColunaNome + 16, CfgRel.PosicaoLinha)
                .DrawString("Telefone", CfgRelatorio.Fontes.Arial7B, Brushes.Black, ColunaTelefone + 0, CfgRel.PosicaoLinha)
                .DrawString("Email", CfgRelatorio.Fontes.Arial7B, Brushes.Black, ColunaEmail + 16, CfgRel.PosicaoLinha)

                CfgRel.PosicaoLinha += 4
                CfgRelatorio.LinhaHorizontal(CfgRelatorio.EnTipoLinha.Fina, CfgRel.MargemEsquerda, CfgRel.MargemDireita, CfgRel.PosicaoLinha, e)
                CfgRel.PosicaoLinha += 2
            End With

        Catch ex As Exception
            Erro.Log(ConexaoSQL.VersaoAplicacao, Assembly.GetExecutingAssembly.GetName.Name(), ex.Message &
                     " " & ex.StackTrace, ToString, MethodBase.GetCurrentMethod().Name, True, Erl.ToString)
        End Try
    End Sub

    Private Sub ImprimeCorpo(e As PrintPageEventArgs)
        Try
            With e.Graphics

                For Ponteiro = Ponteiro To AgendaSelecao.Rows.Count - 1
                    Dim Dado = AgendaSelecao.Rows(Ponteiro)

                    .DrawString(Dado("Codigo"), CfgRelatorio.Fontes.Arial7, Brushes.Black, ColunaCodigo, CfgRel.PosicaoLinha)
                    .DrawString(Dado("Nome"), CfgRelatorio.Fontes.Arial7, Brushes.Black, ColunaNome, CfgRel.PosicaoLinha)
                    .DrawString(Dado("Telefone"), CfgRelatorio.Fontes.Arial7, Brushes.Black, ColunaTelefone, CfgRel.PosicaoLinha)
                    .DrawString(Dado("Email"), CfgRelatorio.Fontes.Arial7, Brushes.Black, ColunaEmail, CfgRel.PosicaoLinha)

                    CfgRel.PosicaoLinha += 4

                    'Mudança de Pagina
                    If CfgRel.PosicaoLinha > (CfgRel.MargemRodape) Then
                        CfgRelatorio.DesenhaRetangulo(e, CfgRelatorio.LinhaFina, CfgRel.MargemEsquerda, CfgRel.MargemTopo,
                                                       CfgRel.MargemDireita - CfgRel.MargemEsquerda, CfgRel.PosicaoLinha - CfgRel.MargemTopo, False)
                        CfgRel.PosicaoLinha += 1
                        .DrawString("Data de Impressão: " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - " & frmLogin.txtUsuario.Text & ConexaoSQL.NomeUsu, CfgRelatorio.Fontes.Arial7, Brushes.Black, CfgRel.MargemEsquerda, CfgRel.PosicaoLinha)

                        CfgRel.PosicaoLinha = CfgRel.MargemTopo
                        CfgRel.NumeroPagina += 1
                        e.HasMorePages = True
                        Ponteiro += 1
                        Exit Sub
                    End If
                Next
                CfgRelatorio.DesenhaRetangulo(e, CfgRelatorio.LinhaFina, CfgRel.MargemEsquerda, CfgRel.MargemTopo,
                                                       CfgRel.MargemDireita - CfgRel.MargemEsquerda, CfgRel.PosicaoLinha - CfgRel.MargemTopo, False)
                CfgRel.PosicaoLinha += 1
                .DrawString("Data de Impressão: " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - " & frmLogin.txtUsuario.Text & ConexaoSQL.NomeUsu, CfgRelatorio.Fontes.Arial7, Brushes.Black, CfgRel.MargemEsquerda, CfgRel.PosicaoLinha)

                e.HasMorePages = False
            End With
        Catch ex As Exception
            Erro.Log(ConexaoSQL.VersaoAplicacao, Assembly.GetExecutingAssembly.GetName.Name(), ex.Message &
                     " " & ex.StackTrace, ToString, MethodBase.GetCurrentMethod().Name, True, Erl.ToString)
        End Try
    End Sub
#End Region
End Class
