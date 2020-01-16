'Written by Daniel Lenshin for Gryphon Robotics (FRC 5549)'

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
Dim autoBPowerCells() As Integer
Dim autoOPowerCells() As Integer
Dim autoIPowerCells() As Integer
Dim autoLineCrossed() As Integer
Dim bPowerCells() As Integer
Dim oPowerCells() As Integer
Dim iPowerCells() As Integer
Dim rotationControl() As Integer
Dim positionControl() As Integer
Dim hangComplete() As Integer
Dim hangBalanced() As Integer
Dim defenseSkill() As Integer
Dim win() As Integer

ReDim teamNumber(DIRows) As Integer
ReDim autoBPowerCells(DIRows) As Integer
ReDim autoOPowerCells(DIRows) As Integer
ReDim autoIPowerCells(DIRows) As Integer
ReDim autoLineCrossed(DIRows) As Integer
ReDim bPowerCells(DIRows) As Integer
ReDim oPowerCells(DIRows) As Integer
ReDim iPowerCells(DIRows) As Integer
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
    autoBPowerCells(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    autoOPowerCells(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    autoIPowerCells(DICell.row - 2) = DICell.Value
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
    bPowerCells(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    oPowerCells(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellCycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)

While Not IsEmpty(DICell)
    iPowerCells(DICell.row - 2) = DICell.Value
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
PDCellCycle = 2

Set PDCell = resetCell(PDCell)
Set PDCell = PDCell.Offset(PDCellCycle, 0)

While Not IsEmpty(PDCell)
    Dim teamNum As Integer
    teamNum = PDCell.Value
    
    Set PDCell = PDCell.Offset(0, 8)
    
    PDCell.Value = getAvg(autoBPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(autoOPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(autoIPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(autoLineCrossed, teamNum, "boolean", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(bPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(bPowerCells, teamNum, "teleop per second", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(oPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(oPowerCells, teamNum, "teleop per second", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(iPowerCells, teamNum, "integer", teamNumber, DIRows)
    Set PDCell = PDCell.Offset(0, 1)
    PDCell.Value = getAvg(iPowerCells, teamNum, "teleop per second", teamNumber, DIRows)
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
