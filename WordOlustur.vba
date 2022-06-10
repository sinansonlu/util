Sub WordOlustur()
	
	'Word kullanabilmek icin obje olustur
	Set WordApp = CreateObject("Word.Application")

	'Belgelerin acik kalmasini istiyorsak asagidaki satir aktif halde olmali
	'WordApp.Visible = True

	For i = 2 To 14
	    'Yeni Word belgesi olustur
	    Set ObjectNewDoc = WordApp.Documents.Add

	    'Belge icerigini doldur
	    With WordApp.Selection
		'Baslik
		.Font.Bold = True
		.Font.Name = "Arial"
		.Font.Size = 16

		'0: sola yasli, 1: ortalanmis, 2: saga yasli
		.ParagraphFormat.Alignment = 1

		.typetext Text:="Test Belgesi"

		'Alt satira gecmek icin
		.TypeParagraph

		'Metin
		.Font.Bold = False
		.Font.Name = "Arial"
		.Font.Size = 12

		.ParagraphFormat.Alignment = 0

		.typetext Text:="Bazı açıklamalar..."	
		.TypeParagraph
		.TypeParagraph

		'Ad
		.ParagraphFormat.Alignment = 0
		.Font.Bold = True
		.Font.Size = 12
		.typetext Text:="Ad: "

		.Font.Bold = False
		.typetext Text:=Cells(i, "A").Value
		.TypeParagraph

		'Miktar
		.Font.Bold = True
		.Font.Size = 12
		.typetext Text:="Miktar: "

		.Font.Bold = False
		.typetext Text:="" & Cells(i, "B").Value
		.TypeParagraph
	    End With

	    'Belgeyi kayit et
	    ObjectNewDoc.SaveAs "D:\Test Belgeleri\Belge_" & i & "_" & Cells(i, "A").Value & ".docx"
	    ObjectNewDoc.Close

	    Set ObjectNewDoc = Nothing

	    'Siradaki satira gec
	    Next i

	Set WordApp = Nothing

End Sub
