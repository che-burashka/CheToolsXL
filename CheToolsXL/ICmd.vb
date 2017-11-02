Imports Office = NetOffice.OfficeApi

Public Interface ICmd
    Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean)
End Interface
