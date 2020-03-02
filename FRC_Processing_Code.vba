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
Dim autoHPowerCells() As Integer
Dim autoLPowerCells() As Integer
Dim autoLineCrossed() As Integer
Dim hPowerCells() As Integer
Dim lPowerCells() As Integer
Dim rotationControl() As Integer
Dim positionControl() As Integer
Dim hangComplete() As Integer
Dim hangBalanced() As Integer
Dim defenseSkill() As Integer
Dim win() As Integer

ReDim teamNumber(DIRows) As Integer
ReDim autoHPowerCells(DIRows) As Integer
ReDim autoLPowerCells(DIRows) As Integer
ReDim autoLineCrossed(DIRows) As Integer
ReDim hPowerCells(DIRows) As Integer
ReDim lPowerCells(DIRows) As Integer
ReDim rotationControl(DIRows) As Integer
ReDim positionControl(DIRows) As Integer
ReDim hangComplete(DIRows) As Integer
ReDim hangBalanced(DIRows) As Integer
ReDim defenseSkill(DIRows) As Integer
ReDim win(DIRows) As Integer

Dim DICellCycle As Integer
DICellCycle = 0

Set DICell = resetCell(DICell)
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    teamNumber(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    autoHPowerCells(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    autoLPowerCells(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    autoLineCrossed(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    hPowerCells(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    lPowerCells(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    rotationControl(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    positionControl(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    hangComplete(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    hangBalanced(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    defenseSkill(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    win(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)

Dim PDCellCycle As Integer
PDCellCycle = 1

Set PDCell = resetCell(PDCell)
Set PDCell = PDCell.Offset(PDCellCycle, 0)

While Not IsEmpty(PDCell)
    Dim teamNum As Integer
    teamNum = PDCell.Value
    
    Set PDCell = PDCell.Offset(0, 9)
    PDCell.Value = getAvg(autoHPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(autoLPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(autoLineCrossed, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(hPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(hPowerCells, teamNum, "teleop per second", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(lPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(lPowerCells, teamNum, "teleop per second", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(rotationControl, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(positionControl, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(hangComplete, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(hangBalanced, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(defenseSkill, teamNum, "quality", teamNumber, DIRows)
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
            getAvg = (sum / amtTeamTerms) / 135
        End If
    
    End If
End Function

'<<Function which returns the cell reset at 1, 1>>'
Function resetCell(rng As Range) As Range

    Set resetCell = Worksheets(rng.Worksheet.Name).Cells(1, 1)
    
End Function

