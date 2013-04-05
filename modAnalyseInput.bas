Attribute VB_Name = "modAnalyseInput"

Public userCommands() As command
Public numCommands As Integer
Public Function keyWords(ByVal inp As String) As Boolean
Dim spaceDel() As String
Dim strTemp As String
Dim temp() As String
spaceDel = Split(inp, " ")
Open App.Path & "\keyDb.txt" For Input As #1
Do Until EOF(1)
    Line Input #1, strTemp
    temp = Split(strTemp, ",")
    For x = LBound(spaceDel) To UBound(spaceDel)
        If temp(0) = spaceDel(x) Then
            ReDim Preserve userCommands(numCommands) As command
            userCommands(numCommands).Text = temp(0)
            userCommands(numCommands).Index = temp(1)
            numCommands = numCommands + 1
        End If
    Next x
Loop
Close #1
End Function

Public Function determineAction() As Boolean
'for every command the keyWords function found, decide what to do with that command. ie. if the commands are "When" and "lunch" what should happen?

'for testing
For x = LBound(userCommands) To UBound(userCommands)
    MsgBox (userCommands(x).Text)
Next x
numCommands = 0
End Function
