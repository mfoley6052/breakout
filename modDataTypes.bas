Attribute VB_Name = "modDataTypes"
Public Type human
    trait(1 To 3) As String
    Size As Integer
    name As String
    gender As String
End Type
Public Type command
    key As Integer
    Text As String
End Type
Public intLines As Integer
Public prisoners() As human
Public numPrisoners As Integer
