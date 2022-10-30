#Region "Imports"
Imports System.Data.SqlClient
Imports xAlimento
Imports GGeral
#End Region

Public Class frmMDIAlpha

#Region "Única Instância"
    Private Shared _Instance As frmMDIAlpha = Nothing
    Private Shared _NovaInstancia As Boolean = False

    Public Property NovaInstancia() As String
        Get
            Return _NovaInstancia
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Shared Function Instance() As frmMDIAlpha
        If _Instance Is Nothing OrElse _Instance.IsDisposed = True Then
            _Instance = New frmMDIAlpha
            _NovaInstancia = True
        Else
            _NovaInstancia = False
        End If
        _Instance.BringToFront()
        Return _Instance
    End Function
#End Region

    Private Sub frmMDIAlpha_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmLogin.ShowDialog()
    End Sub

    Private Sub mnuClientesInclui_Click(sender As Object, e As EventArgs) Handles mnuClientesIncluir.Click
        Dim myfrmAgenda As frmAgenda

        myfrmAgenda = frmAgenda.Instance
        myfrmAgenda.TipoForm = frmAgenda.enumTipoForm.Inclui
        myfrmAgenda.MdiParent = Me
        myfrmAgenda.Show()
    End Sub

    Private Sub mnuClientesSelecionar_Click(sender As Object, e As EventArgs) Handles mnuClientesSelecionar.Click
        Dim myfrmSeleciona As frmSeleciona

        myfrmSeleciona = frmSeleciona.Instance
        myfrmSeleciona.MdiParent = Me
        myfrmSeleciona.Show()
    End Sub

    Private Sub mnuCategoria_Click(sender As Object, e As EventArgs) Handles mnuCategoria.Click
        Dim myfrmCategoria As frmCategoria

        myfrmCategoria = frmCategoria
        myfrmCategoria.MdiParent = Me
        myfrmCategoria.Show()
    End Sub

    Private Sub mnuRelatorio_Click(sender As Object, e As EventArgs) Handles mnuRelatorio.Click
        Dim myfrmRelatorio As frmRelatorio

        myfrmRelatorio = frmRelatorio
        myfrmRelatorio.MdiParent = Me
        myfrmRelatorio.Show()
    End Sub
    Private Sub mnuIncluiPedido_Click(sender As Object, e As EventArgs) Handles mnuPedidoIncluir.Click
        Dim myfrmIncluiPedido As frmIncluiPedido

        myfrmIncluiPedido = frmIncluiPedido
        myfrmIncluiPedido = frmIncluiPedido.Instance
        myfrmIncluiPedido.MdiParent = Me
        myfrmIncluiPedido.Show()
    End Sub

    Private Sub mnuPedidoSelecionar_Click(sender As Object, e As EventArgs) Handles mnuPedidoSelecionar.Click
        Dim myfrmSelecionaPedido As frmSelecionaPedido

        myfrmSelecionaPedido = frmSelecionaPedido
        myfrmSelecionaPedido = frmSelecionaPedido.Instance
        myfrmSelecionaPedido.MdiParent = Me
        myfrmSelecionaPedido.Show()
    End Sub
End Class