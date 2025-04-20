Attribute VB_Name = "modMapping"
Option Compare Database
Option Explicit

Function OpenMap(Address, City, State, Zip, Country)

    Dim strAddress As String
    strAddress = Nz(Address)
    strAddress = strAddress & IIf(strAddress = "", "", ", ") & Nz(City)
    strAddress = strAddress & IIf(strAddress = "", "", ", ") & Nz(State)
    strAddress = strAddress & IIf(strAddress = "", "", ", ") & Nz(Zip)
    strAddress = strAddress & IIf(strAddress = "", "", ", ") & Nz(Country)
    
    If strAddress = "" Then
        MsgBox "There is no address to map."
    Else
        Application.FollowHyperlink "http://maps.live.com/default.aspx?where1=" & strAddress
    End If
End Function
