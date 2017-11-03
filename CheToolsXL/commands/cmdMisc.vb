Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi

Imports ExcelConsts = NetOffice.ExcelApi.Enums

Public Class Utils

    Shared Sub NormalizeWbk(ByVal wbk As Excel.Workbook)
        '
        Dim sht As Excel.Worksheet
        Dim win As Excel.Window

        Try
            For Each sht In wbk.Worksheets
                sht.DisplayPageBreaks = False
            Next
            For Each win In wbk.Windows
                win.DisplayGridlines = False
                win.Zoom = 100
            Next
        Catch ex As Exception
            'MsgBox(Err.Description)
        End Try


    End Sub

    Shared Sub NormalizeAll(ByVal xl As Excel.Application)

        Dim wbk As Excel.Workbook

        xl.DisplayAlerts = False
        xl.Calculation = Excel.Enums.XlCalculation.xlCalculationManual
        xl.CalculateBeforeSave = False
        'xl.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        xl.DisplayPasteOptions = False
        xl.FixedDecimalPlaces = 2
        xl.SheetsInNewWorkbook = 1
        'xl.ShowWindowsInTaskbar = False

        For Each wbk In xl.Workbooks
            Utils.NormalizeWbk(wbk)
        Next

    End Sub

End Class


Public Class cmdSetMySettings
    '' normalize everything
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
        'Utils.NormalizeAll(m_host)
    End Sub
End Class


Public Class cmdHighlightNamedRanges
    '' normalize everything
    Implements ICmd
    Private WithEvents m_btn As Office.CommandBarButton
    Private m_host As Excel.Application
    Private m_doHighlight As Boolean

    Public Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean) Implements ICmd.Init
        m_btn = btn
        m_host = host
        m_doHighlight = fwd
    End Sub

    Private Sub clickHandler(ByVal b As Office.CommandBarButton, ByRef cd As Boolean) Handles m_btn.ClickEvent
        Run()
    End Sub

    Sub Run()
        Dim nm As Excel.Name
        For Each nm In m_host.ActiveWorkbook.Names
            Try
                If m_doHighlight Then
                    highlightName(nm, 3)
                Else
                    highlightName(nm, -1)
                End If
            Catch e As Exception
                '
            End Try
        Next
    End Sub

    Private Sub highlightName(ByVal nm As Excel.Name, ByVal clr As Long)
        '

        Dim r As Excel.Range

        If Strings.InStr(nm.RefersTo, "#") <> 0 Then
            nm.Delete()
        End If
        r = nm.RefersToRange
        r.Cells(1, 1).ClearComments()



        r.Borders(ExcelConsts.XlBordersIndex.xlDiagonalDown).LineStyle = ExcelConsts.XlLineStyle.xlLineStyleNone
        r.Borders(ExcelConsts.XlBordersIndex.xlDiagonalUp).LineStyle = ExcelConsts.XlLineStyle.xlLineStyleNone
        r.Borders(ExcelConsts.XlBordersIndex.xlEdgeLeft).LineStyle = ExcelConsts.XlLineStyle.xlLineStyleNone
        r.Borders(ExcelConsts.XlBordersIndex.xlEdgeTop).LineStyle = ExcelConsts.XlLineStyle.xlLineStyleNone
        r.Borders(ExcelConsts.XlBordersIndex.xlEdgeBottom).LineStyle = ExcelConsts.XlLineStyle.xlLineStyleNone
        r.Borders(ExcelConsts.XlBordersIndex.xlEdgeRight).LineStyle = ExcelConsts.XlLineStyle.xlLineStyleNone
        r.Borders(ExcelConsts.XlBordersIndex.xlInsideVertical).LineStyle = ExcelConsts.XlLineStyle.xlLineStyleNone
        r.Borders(ExcelConsts.XlBordersIndex.xlInsideHorizontal).LineStyle = ExcelConsts.XlLineStyle.xlLineStyleNone

        If clr >= 0 Then
            r.BorderAround()
            r.BorderAround(ExcelConsts.XlLineStyle.xlLineStyleNone, ExcelConsts.XlBorderWeight.xlThin, clr)
            Dim tmp As String
            tmp = nm.Name
            r.Cells(1, 1).AddComment(tmp)
        Else
            ''
        End If

    End Sub
End Class