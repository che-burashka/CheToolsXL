Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi
Imports System.Windows.Forms

Public Class cmdPasteSpecial


    Implements ICmd
    Private WithEvents m_btn As Office.CommandBarButton
    Private m_host As Excel.Application

    Public Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean) Implements ICmd.Init
        m_btn = btn
        m_host = host
    End Sub

    Private Sub clickHandler(ByVal b As Office.CommandBarButton, ByRef cd As Boolean) Handles m_btn.ClickEvent
        Try
            doPaste()
        Catch
            ''
        End Try
    End Sub


    Private Sub doPaste()

        Dim sel As Excel.Range

        Dim trans As Integer
        Dim vals As Integer
        Dim frmls As Integer
        Dim fmts As Integer
        Dim clearAfter As Integer
        'clearAfter = frmPasteSpecial.DefInstance.chkClearClipboard.CheckState

        If Not (TypeOf (m_host.Selection) Is Excel.Range) Then
            ''If TypeName(m_host.Selection) <> "Range" Then
            MsgBox("Unsupported selection type")
            Exit Sub
        End If

        '''frmPasteSpecial.DefInstance.ShowDialog()

        If False Then 'frmPasteSpecial.DefInstance.OKCancel Then

            On Error GoTo tryagain

            sel = m_host.Selection

            'trans = frmPasteSpecial.DefInstance.chkTranspose.CheckState
            'vals = frmPasteSpecial.DefInstance.chkValues.CheckState
            'frmls = frmPasteSpecial.DefInstance.chkFormulas.CheckState
            'fmts = frmPasteSpecial.DefInstance.chkFormats.CheckState

            If vals = 1 Then
                If trans = 1 Then sel.PasteSpecial(Excel.Enums.XlPasteType.xlPasteValues, Excel.Enums.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, True)
                If trans = 0 Then sel.PasteSpecial(Excel.Enums.XlPasteType.xlPasteValues, Excel.Enums.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, False)
            End If

            If frmls = 1 Then
                If trans = 1 Then sel.PasteSpecial(Excel.Enums.XlPasteType.xlPasteFormulas, Excel.Enums.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, True)
                If trans = 0 Then sel.PasteSpecial(Excel.Enums.XlPasteType.xlPasteFormulas, Excel.Enums.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, False)
            End If

            If fmts = 1 Then
                If trans = 1 Then sel.PasteSpecial(Excel.Enums.XlPasteType.xlPasteFormats, Excel.Enums.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, True)
                If trans = 0 Then sel.PasteSpecial(Excel.Enums.XlPasteType.xlPasteFormats, Excel.Enums.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, False)
            End If

            If clearAfter = 1 Then
                Clipboard.Clear()
            End If

            Exit Sub

        End If

tryagain:

        If MsgBox("Ooops, no go. Try again as text?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            'UPGRADE_WARNING: Couldn't resolve default property of object oHostApp.ActiveSheet.PasteSpecial. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            m_host.ActiveSheet.PasteSpecial(Format:="Csv", Link:=False, DisplayAsIcon:=False)
        End If

    End Sub


End Class