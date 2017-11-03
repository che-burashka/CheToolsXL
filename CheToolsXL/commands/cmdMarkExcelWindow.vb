Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi

Public Class cmdMarkExcelWindow

    Implements ICmd
    Private WithEvents m_btn As Office.CommandBarButton
    Private m_host As Excel.Application

    Public Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean) Implements ICmd.Init
        m_btn = btn
        m_host = host
    End Sub

    Private Sub clickHandler(ByVal b As Office.CommandBarButton, ByRef cd As Boolean) Handles m_btn.ClickEvent
        frmApplicationCaption.oHostApp = m_host
        frmApplicationCaption.DefInstance.ShowDialog()
    End Sub


End Class
