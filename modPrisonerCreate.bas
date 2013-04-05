Attribute VB_Name = "modPrisonerCreate"
Public Function newPrisoner(Optional ByVal numToCreate As Integer) As Boolean
Randomize Timer
Dim rand As Integer
Dim lineNum As Long
Dim lenFile As Long
Dim curLine As Long
Dim temp() As String
Close #1
If numToCreate = 0 Then
    numToCreate = 1
End If
For z = 1 To numToCreate
    numPrisoners = numPrisoners + 1
    Open App.Path & "\traits.txt" For Input As #1
    ReDim Preserve prisoners(numPrisoners) As human
    For x = 1 To 3
    lenFile = 0
        Do While Not EOF(1)
            Line Input #1, prisoners(numPrisoners).trait(x)
            lenFile = lenFile + 1
        Loop
        lineNum = Int((Rnd() * lenFile) + 1)
        Seek #1, 1
        curLine = 1
         
        Do While curLine <= lineNum
            Line Input #1, prisoners(numPrisoners).trait(x)
            curLine = curLine + 1
        Loop
    Next x
    Close #1
    Open App.Path & "\yob1990.txt" For Input As #1
    lenFile = 0
        Do While Not EOF(1)
            Line Input #1, prisoners(numPrisoners).name
         temp() = Split(prisoners(numPrisoners).name, " ")
         prisoners(numPrisoners).name = temp(0)
         prisoners(numPrisoners).gender = temp(1)
            lenFile = lenFile + 1
        Loop
        lineNum = Int((Rnd() * lenFile) + 1)
        Seek #1, 1
        curLine = 1
         
        Do While curLine <= lineNum
            Line Input #1, prisoners(numPrisoners).name
            temp() = Split(prisoners(numPrisoners).name, " ")
            prisoners(numPrisoners).name = temp(0)
            prisoners(numPrisoners).gender = temp(1)
            curLine = curLine + 1
        Loop
    Close #1
    Call eventOccur("New Prisoner arrived: " & prisoners(numPrisoners).name)
    Form1.lstPrisoners.AddItem (prisoners(numPrisoners).name & "    " & prisoners(numPrisoners).trait(1) & "    " & prisoners(numPrisoners).trait(2) & "    " & prisoners(numPrisoners).trait(3))
Next z
End Function
