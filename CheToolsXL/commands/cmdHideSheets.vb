Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi

Public Class cmdHideSheets

    Implements ICmd
    Private WithEvents m_btn As Office.CommandBarButton
    Private m_host As Excel.Application
    Private m_doHide As Boolean

    Public Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean) Implements ICmd.Init
        m_btn = btn
        m_host = host
        m_doHide = fwd
    End Sub

    Private Sub clickHandler(ByVal b As Office.CommandBarButton, ByRef cd As Boolean) Handles m_btn.ClickEvent
        If m_doHide Then
            Hide()
        Else
            Unhide()
        End If
    End Sub


    Sub Hide()

        Dim wbk As Excel.Workbook
        Dim sht As Excel.Worksheet

        wbk = m_host.ActiveWorkbook
        Dim nm As String

        For Each sht In wbk.Worksheets
            nm = sht.Name
            If Strings.Left(nm, 1) = "_" Then
                sht.Visible = Excel.Enums.XlSheetVisibility.xlSheetHidden
            End If
        Next


    End Sub

    Sub Unhide()

        Dim wbk As Excel.Workbook
        Dim sht As Excel.Worksheet

        wbk = m_host.ActiveWorkbook

        For Each sht In wbk.Worksheets
            sht.Visible = Excel.Enums.XlSheetVisibility.xlSheetVisible
        Next

    End Sub
End Class