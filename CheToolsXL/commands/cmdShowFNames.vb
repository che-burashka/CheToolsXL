Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi

Public Class cmdShowFNames

    Implements ICmd
    Private WithEvents m_btn As Office.CommandBarButton
    Private m_host As Excel.Application

    Public Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean) Implements ICmd.Init
        m_btn = btn
        m_host = host
    End Sub

    Private Sub clickHandler(ByVal b As Office.CommandBarButton, ByRef cd As Boolean) Handles m_btn.ClickEvent

        Dim wbk As Excel.Workbook
        Dim title As String
        Dim path As String
        Dim w As Excel.Window
        Dim i As Integer

        For Each wbk In m_host.Workbooks

            title = wbk.Name
            path = wbk.Path

            If wbk.ReadOnly Then
                title = title & " [Read Only]"
            End If

            title = title & "  (" & path & ")"

            For i = 1 To wbk.Windows.Count
                w = wbk.Windows.Item(i)
                w.Caption = title
            Next

        Next
    End Sub

End Class
