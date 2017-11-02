Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports Office = NetOffice.OfficeApi

Public Class cmdCalcSelection

    Implements ICmd
    Private WithEvents m_btn As Office.CommandBarButton
    Private m_host As Excel.Application

    Public Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean) Implements ICmd.Init
        m_btn = btn
        m_host = host
    End Sub

    Private Sub clickHandler(ByVal b As Office.CommandBarButton, ByRef cd As Boolean) Handles m_btn.ClickEvent
        Run()
    End Sub


    Sub Run()
        Dim selObj As Object
        selObj = m_host.Selection
        Dim selRge As Excel.Range
        If TypeOf (selObj) Is Excel.Range Then
            selRge = selObj
            selRge.Calculate()
        End If
    End Sub


End Class