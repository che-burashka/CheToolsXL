Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi

Public Class cmdFreeze

    Implements ICmd
    Private WithEvents m_btn As Office.CommandBarButton
    Private m_host As Excel.Application
    Private m_DoFreeze As Boolean

    Public Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean) Implements ICmd.Init

        m_btn = btn
        m_host = host
        m_DoFreeze = fwd

    End Sub

    Private Sub clickHandler(ByVal b As Office.CommandBarButton, ByRef cd As Boolean) Handles m_btn.ClickEvent

        If m_DoFreeze Then
            DoFreeze()
        Else
            DoThaw()
        End If

    End Sub

    Private Function hasFormulaComment(ByVal rge As Excel.Range) As Boolean

        hasFormulaComment = False
        Dim commentString As String
        commentString = ""

        Try
            commentString = rge.Comment.Text
            If commentString.Chars(0) = "=" Then
                hasFormulaComment = True
            End If
        Catch ex As Exception
            '
        End Try

    End Function



    Sub DoFreeze()

        Dim selObj As Object
        selObj = m_host.Selection
        Dim selRge As Excel.Range
        Dim c1 As Excel.Range
        Dim tmp As Object

        If TypeOf (selObj) Is Excel.Range Then
            selRge = selObj

            For Each c1 In selRge.Cells
                If c1.HasFormula Then
                    tmp = c1.Value
                    c1.AddComment(c1.Formula)
                    c1.Formula = ""
                    c1.Value = tmp
                End If
            Next
        End If

    End Sub

    Sub DoThaw()

        Dim selObj As Object
        selObj = m_host.Selection
        Dim selRge As Excel.Range
        Dim c1 As Excel.Range

        If TypeOf (selObj) Is Excel.Range Then

            selRge = selObj
            For Each c1 In selRge.Cells
                If hasFormulaComment(c1) Then
                    c1.Formula = c1.Comment.Text
                    c1.ClearComments()
                End If
            Next
        End If

    End Sub


End Class