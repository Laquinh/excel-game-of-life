Attribute VB_Name = "ModuleSetup"
Sub CreateSheet()
	Dim sheet as Worksheet
	On Error Resume Next
	Set sheet = Sheets("GameOfLife")
	If Not sheet Is Nothing Then
		Application.DisplayAlerts = False
		sheet.Delete
		Application.DisplayAlerts = True
	End If
	Sheets.Add.Name = "GameOfLife"
End Sub

Sub InitSettings()
	Range("B2").Value = "Width"
	Range("B3").Value = "Height"
	Range("B4").Value = "Top-left cell"
	Range("B5").Value = "Iterations"
	With Range("B2:B5")
		.Font.Bold = True
		.ColumnWidth = .ColumnWidth*2
	End With

	Range("C2").Value = 50
	Range("C3").Value = 50
	Range("C4").Value = "G8"
	Range("C5").Value = 10
	With Range("C2:C5")
        .HorizontalAlignment = xlRight
	End With
End Sub

Sub CreateButtons()
	Dim rectangle As Range

	Set rectangle = ActiveSheet.Range(Cells(2, 5), Cells(2, 5))
	ActiveSheet.Buttons.Add(rectangle.Left, rectangle.Top, rectangle.Width, rectangle.Height).Select
	Selection.OnAction = "Clear"
	Selection.Characters.Text = "Clear"

	Set rectangle = ActiveSheet.Range(Cells(4, 5), Cells(4, 5))
	ActiveSheet.Buttons.Add(rectangle.Left, rectangle.Top, rectangle.Width, rectangle.Height).Select
	Selection.OnAction = "CreateTable"
	Selection.Characters.Text = "Create Table"

	Set rectangle = ActiveSheet.Range(Cells(6, 5), Cells(6, 5))
	ActiveSheet.Buttons.Add(rectangle.Left, rectangle.Top, rectangle.Width, rectangle.Height).Select
	Selection.OnAction = "Run"
	Selection.Characters.Text = "Run"

	Cells(2, 5).ColumnWidth = Cells(2, 5).ColumnWidth*2
End Sub

Sub Clear()
	Dim gameRange As Range
	Set gameRange = Range(Range("C4") & ":INDEX(1048576:1048576,Columns(1:1))")
	gameRange.Clear
End Sub

Sub InitTableBorders(gameRange As Range)
	With gameRange
		.Borders(xlDiagonalDown).LineStyle = xlNone
    	.Borders(xlDiagonalUp).LineStyle = xlNone
		.Borders(xlInsideVertical).LineStyle = xlNone
		.Borders(xlInsideHorizontal).LineStyle = xlNone
	End With
    
    With gameRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With gameRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With gameRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With gameRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End Sub

Sub InitTableFormat(gameRange As Range)
	With gameRange
		.ColumnWidth = 1
		.RowHeight = 9.75
		.Value = "0"
		.NumberFormat = ";;;"
	End With

    gameRange.FormatConditions.AddColorScale ColorScaleType:=2
    gameRange.FormatConditions(gameRange.FormatConditions.Count).SetFirstPriority
    gameRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    gameRange.FormatConditions(1).ColorScaleCriteria(1).Value = 0
    With gameRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    gameRange.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    gameRange.FormatConditions(1).ColorScaleCriteria(2).Value = 1
    With gameRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Sub

Function InitGameRange() As Range
	Dim topLeft As Range
	Dim bottomRight As Range
	Dim width As Integer
	Dim height As Integer

	width = Range("C2").Value
	height = Range("C3").Value

	Set topLeft = Range(Range("C4"))
	Set bottomRight = topLeft.Offset(height - 1, width - 1)

	Set InitGameRange = Range(topLeft, bottomRight)
End Function

Sub CreateTable()
	Dim gameRange As Range
	Set gameRange = InitGameRange()

	Call InitTableBorders(gameRange)
	Call InitTableFormat(gameRange)
End Sub

Sub Run()
	MsgBox "Run"
End Sub

Sub Setup()
	CreateSheet
	InitSettings
	CreateButtons
	Range("A1").Select
End Sub
