Attribute VB_Name = "modUpdateEvents"
Option Explicit
Public x As Integer
Public Function eventOccur(ByVal eventMsg As String) As Boolean
With Form1
    .txtUpdates(4).Text = .txtUpdates(3).Text
    .txtUpdates(3).Text = .txtUpdates(2).Text
    .txtUpdates(2).Text = .txtUpdates(1).Text
    .txtUpdates(1).Text = .txtUpdates(0).Text
    .txtUpdates(0).Text = eventMsg
End With
End Function


