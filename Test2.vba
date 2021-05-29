Sub Test2()
	Dim d As Object
	Dim d1 As Object
	Dim d2 As Object

	Set d = CreateObject("Scripting.Dictionary")
	Set d1 = CreateObject("Scripting.Dictionary")
	Set d2 = CreateObject("Scripting.Dictionary")

	Last = Worksheets("ALL").Range("A" & Rows.Count).End(xlUp).Row

	For i = 2 To Last
	  d(CSng(Worksheets("ALL").Cells(i, 5))) = 0
	  d1(CSng(Worksheets("ALL").Cells(i, 5))) = 0
	  d2(CSng(Worksheets("ALL").Cells(i, 5))) = 0
	Next i

	For i = 2 To Last
	  d(CSng(Worksheets("ALL").Cells(i, 5))) = d(CSng(Worksheets("ALL").Cells(i, 5))) + 1
	  d1(CSng(Worksheets("ALL").Cells(i, 5))) = d1(CSng(Worksheets("ALL").Cells(i, 5))) + CSng(Worksheets("ALL").Cells(i, 7))
	  
	  If CSng(Worksheets("ALL").Cells(i, 7)) > 70 Then
			d2(CSng(Worksheets("ALL").Cells(i, 5))) = d2(CSng(Worksheets("ALL").Cells(i, 5))) + 1
	  End If
	  
	Next i

	Worksheets("Toplamlar").Cells(1, 1) = "Age"
	Worksheets("Toplamlar").Cells(1, 2) = "Count"
	Worksheets("Toplamlar").Cells(1, 3) = "Total"
	Worksheets("Toplamlar").Cells(1, 4) = "X > 70"

	Worksheets("Toplamlar").Range("A2").Resize(d.Count) = Application.Transpose(d.keys)
	Worksheets("Toplamlar").Range("B2").Resize(d.Count) = Application.Transpose(d.items)
	Worksheets("Toplamlar").Range("C2").Resize(d1.Count) = Application.Transpose(d1.items)
	Worksheets("Toplamlar").Range("D2").Resize(d2.Count) = Application.Transpose(d2.items)
End Sub
