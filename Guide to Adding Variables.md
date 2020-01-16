Guide to Adding your own Variables

Adding your own variables to the spreadsheet which your team wishes to keep track of is possible, but will require modification of the code.
It will require at least some experience with VBA for Excel.

First, add a column to the Processed Data spreadsheet where your variable will be displayed, and do the same on the Data Input spreadsheet.
Make sure that both of the added columns are in the same place relative to the other variables (e.g. between the same variables)

Declare an array with the name of your variable **As Integer.** Then, redimension the array with the value DIRows, in order to 
resize the array to the amount of rows there are in the "Data Input" spreadsheet. 

Starting at around line 60, there are code blocks that look ike this:

```
While Not IsEmpty(DICell)
    ARRAY(DICell.row - 2) = DICell.Value
    Set DICell = DICell.Offset(1, 0)
Wend
Set DICell = resetCell(DICell)
DICellSycle = DICellCycle + 1
Set DICell = DICell.Offset(2, DICellCycle)
```

where ARRAY() is the array which is being modified in the code. Add another code block line this, making sure that ARRAY() has the same name
as the array you have added. **Make sure to add this block in the right position relative to the columns in the Data Input spreadsheet.**

Starting at around line 180, there are several code blocks which look like this:

```
PDCell.Value = getAvg(ARRAY, teamNum, "TYPE", teamNumber, DIRows)
Set PDCell = PDCell.Offset(0, 1)
```

Add another code block like this, making sure that it is in the correct position. ARRAY is replaced with the name of your new variable, 
and "TYPE" is replaced with the type of average you wish to display. The types are listed below:

**"integer" or "boolean"** - The spreadsheet will display these values as a numerical average. Boolean values are entered as 0 or 1, therefore
the displayed average will be a percentage.

**"quality"** - The spreadsheet will display a numerical average, excluding the terms which are equal to zero. This is meant to measure
how well a team does something, where 0 represents the team not doing anything. Note that the current message for "no quality above 0"
is "NO DEF," but this can be changed on line 269.

**"games played"** - The spreadsheet will display how many terms there are of this data point for the team. 

**"teleop per second"** - The spreadsheet will display an integer value divided by 135, which is the length of time that the tele-operated period lasts.
