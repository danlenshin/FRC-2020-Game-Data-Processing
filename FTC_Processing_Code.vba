'<<Written by Daniel Lenshin>>'

'<<Main Processor Code>>'

Sub Processor()

Dim PDCell As Range
Dim DICell As Range

Set PDCell = Worksheets("Processed Data").Cells(1, 1)
Set DICell = Worksheets("Data Input").Cells(1, 1)

Set DICell = DICell.Offset(2, 0)

Dim DIRows As Integer
DIRows = 0

While Not IsEmpty(DICell)
    DIRows = DIRows + 1
    Set DICell = DICell.Offset(1, 0)
Wend

Set DICell = resetCell(DICell)

Dim teamNumber() As Integer
Dim autoPark() As Integer
Dim blocksDelivered() As Integer
Dim skybridgeCrossed() As Integer
Dim blocks() As Integer
Dim skyscraperLevel() As Integer
Dim cap() As Integer
Dim foundationMoved() As Integer
Dim park() As Integer
Dim win() As Integer

ReDim teamNumber(DIRows) As Integer
ReDim autoPark(DIRows) As Integer
ReDim blocksDelivered(DIRows) As Integer
ReDim skybridgeCrossed(DIRows) As Integer
ReDim blocks(DIRows) As Integer
ReDim skyscraperLevel(DIRows) As Integer
ReDim cap(DIRows) As Integer
ReDim foundationMoved(DIRows) As Integer
ReDim park(DIRows) As Integer
ReDim win(DIRows) As Integer

Dim DICellCycle As Integer
DICellCycle = 0

Set DICell = resetCell(DICell)
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    teamNumber(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    autoPark(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    blocksDelivered(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    skybridgeCrossed(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    blocks(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    skyscraperLevel(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    cap(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    foundationMoved(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    park(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    win(DICell.Row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

Dim PDCellCycle As Integer
PDCellCycle = 2

Set PDCell = resetCell(PDCell)
Set PDCell = PDCell.Offset(PDCellCycle, 0)

While Not IsEmpty(PDCell)
    Dim teamNum As Integer
    teamNum = PDCell.Value
    
    Set PDCell = PDCell.Offset(0, 7)
    PDCell.Value = getAvg(autoPark, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(blocksDelivered, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(skybridgeCrossed, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(blocks, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(blocks, teamNum, "teleop per second", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(skyscraperLevel, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(cap, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(foundationMoved, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(park, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(win, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(win, teamNum, "games played", teamNumber, DIRows)
    
    Set PDCell = resetCell(PDCell)
    PDCellCycle = PDCellCycle + 1
    Set PDCell = PDCell.Offset(PDCellCycle, 0)
Wend

End Sub

'<<Function to get the average of the values of an array corresponding to a certain team>>'
Function getAvg(arr() As Integer, team As Integer, varType As String, teamNumArr() As Integer, DIR As Integer) As Variant

    Dim sum As Integer
    Dim qualAmtZero As Integer
    Dim amtTeamTerms As Integer '< To find the amount of columns which are about the team (to resize equalTerms() array)
    
    sum = 0
    qualAmtZero = 0
    amtTeamTerms = 0
    
    For i = 0 To DIR
        If teamNumArr(i) = team Then
            amtTeamTerms = amtTeamTerms + 1
        End If
    Next i
    
    Dim equalTerms() As Integer
    ReDim equalTerms(amtTeamTerms)
    
    Dim etPusher As Integer
    etPusher = 0
    
    For i = 0 To DIR
        If teamNumArr(i) = team Then
            equalTerms(etPusher) = arr(i)
            etPusher = etPusher + 1
        End If
    Next i
    
    
    If varType = "integer" Or varType = "boolean" Then
        For i = 0 To amtTeamTerms
            sum = sum + equalTerms(i)
        Next i
        
        If amtTeamTerms = 0 Then
            getAvg = ""
        Else
            getAvg = sum / amtTeamTerms
        End If
    ElseIf varType = "quality" Then
        For i = 0 To amtTeamTerms - 1
            If equalTerms(i) = 0 Then
                qualAmtZero = qualAmtZero + 1
            Else
                sum = sum + equalTerms(i)
            End If
        Next i
        
        If amtTeamTerms = 0 Then
            getAvg = ""
        ElseIf amtTeamTerms = qualAmtZero Then
            getAvg = "NO DEF"
        Else
            getAvg = sum / (amtTeamTerms - qualAmtZero)
        End If
    
    ElseIf varType = "games played" Then
        getAvg = amtTeamTerms
        
    ElseIf varType = "teleop per second" Then
                For i = 0 To amtTeamTerms
            sum = sum + equalTerms(i)
        Next i
        
        If amtTeamTerms = 0 Then
            getAvg = ""
        Else
            getAvg = (sum / amtTeamTerms) / 120
        End If
    
    End If
End Function

'<<Function which returns the cell reset at 1, 1>>'
Function resetCell(rng As Range) As Range

    Set resetCell = Worksheets(rng.Worksheet.Name).Cells(1, 1)
    
End Function
