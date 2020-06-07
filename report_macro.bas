Attribute VB_Name = "NewMacros"
Sub ReportArrange()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim par_n As Long
    par_n = doc.Paragraphs.Count
    Dim time As Date
    time = Date
    Dim fime_s As String
    time_s = Year(time) & "”N" & Month(time) & "ŒŽ" & Day(time) & "“úì¬"
    
    Dim title As String
    title = InputBox(Prompt, "ƒ^ƒCƒgƒ‹‚ð“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B")
    
    Dim fir_info As Range
    Set fir_info = doc.Range(Start:=0, End:=0)
    With fir_info
        .InsertAfter Text:=title
        .InsertParagraphAfter
        With .Font
            .Bold = False
            .Size = 14
            .Name = "‚l‚r ‚oƒSƒVƒbƒN"
            .Name = "Arial"
        End With
    End With
    Dim sec_info As Range
    Set sec_info = doc.Range(Start:=doc.Paragraphs(1).Range.End, End:=doc.Paragraphs(1).Range.End)
    With sec_info
        .InsertAfter Text:="(–¼‘O)"
        .InsertParagraphAfter
        .InsertAfter Text:="(Š‘®) (Šw¶”Ô†)"
        .InsertParagraphAfter
        .InsertAfter Text:=time_s
        .InsertParagraphAfter
        With .Font
            .Bold = False
            .Size = 12
            .Name = "‚l‚r ‚oƒSƒVƒbƒN"
            .Name = "Arial"
        End With
        Dim writer_range1 As Range, writer_range2 As Range
        Set writer_range1 = doc.Range(Start:=doc.Paragraphs(1).Range.Start, End:=doc.Paragraphs(1).Range.End)
        Set writer_range2 = doc.Range(Start:=doc.Paragraphs(2).Range.Start, End:=doc.Paragraphs(4).Range.End)
        writer_range1.Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        writer_range2.Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    
    With doc.PageSetup
        .TopMargin = MillimetersToPoints(25)
        .BottomMargin = MillimetersToPoints(25)
        .LeftMargin = MillimetersToPoints(25)
        .RightMargin = MillimetersToPoints(25)
        .TextColumns.SetCount 1
        .CharsLine = 46
        .LinesPage = 42
    End With
    
    Dim counter As Long
    For counter = 5 To (par_n + 4)
        Dim par As Range
        Set par = doc.Range(Start:=doc.Paragraphs(counter).Range.Start, End:=doc.Paragraphs(counter).Range.End)
        par.Select
        Dim contents As String
        contents = par.Text
        Dim fir_letter As String
        fir_letter = Left(contents, 1)
        If fir_letter = "#" Then
            par.SetRange Start:=par.Start, End:=par.Start + 1
            par.Delete
            par.SetRange Start:=doc.Paragraphs(counter).Range.Start, End:=doc.Paragraphs(counter).Range.End
            With par.Font
                .Bold = False
                .Size = 12
                .Name = "‚l‚r ‚o–¾’©"
                .Name = "Times New Roman"
            End With
        Else
        If fir_letter = "%" Then
            par.SetRange Start:=par.Start, End:=par.Start + 1
            par.Delete
            par.SetRange Start:=doc.Paragraphs(counter).Range.Start, End:=doc.Paragraphs(counter).Range.End
            With par.Font
                .Bold = False
                .Size = 9
                .Name = "‚l‚r ‚o–¾’©"
                .Name = "Times New Roman"
            End With
        Else
        If fir_letter = ">" Then
            par.SetRange Start:=par.Start, End:=par.Start + 1
            par.Delete
            par.SetRange Start:=doc.Paragraphs(counter).Range.Start, End:=doc.Paragraphs(counter).Range.End
            Selection.Paragraphs.CharacterUnitLeftIndent = 1
            With par.Font
                .Bold = False
                .Italic = True
                .Size = 10.5
                .Name = "‚l‚r ‚o–¾’©"
                .Name = "Times New Roman"
            End With
        Else
            With par.Font
                .Bold = False
                .Size = 10.5
                .Name = "‚l‚r ‚o–¾’©"
                .Name = "Times New Roman"
            End With
        End If
        End If
        End If
    Next
End Sub
