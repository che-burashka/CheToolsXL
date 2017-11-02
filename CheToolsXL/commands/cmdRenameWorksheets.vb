Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi

Public Class cmdRenameWorksheets

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
        Dim prefix As String

        Dim nm As String
        Dim wbk As Excel.Workbook

        Dim n As Long
        n = 1000

        If TypeOf (selObj) Is Excel.Range Then
            prefix = selObj.Value
            wbk = m_host.ActiveWorkbook
            For Each sht As Excel.Worksheet In wbk.Worksheets
                nm = sht.Name
                sht.Range("A1").ClearComments()
                sht.Range("A1").AddComment(nm)
                n = n + 1
                nm = prefix & "." & n.ToString()
                sht.Name = nm
            Next
        End If


    End Sub


End Class