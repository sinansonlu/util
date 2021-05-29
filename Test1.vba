Sub Test1()
	Last = Worksheets("Gruplar").Range("E" & Rows.Count).End(xlUp).Row

	For i = 2 To Last
		ID = Worksheets("Gruplar").Cells(i, "D").Value
		GRADE = Worksheets("Gruplar").Cells(i, "E").Value
		
		Set FoundCell = Worksheets("ALL").Range("A:A").Find(What:=ID)
		
		If Not FoundCell Is Nothing Then
			Worksheets("ALL").Cells(FoundCell.Row, "G").Value = GRADE
		Else
			MsgBox ("Cannot find " & ID)
		End If
		
		Next i
End Sub
