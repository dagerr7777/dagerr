Sub Publico_xml_sport_foto_filename()

 ' usuwanie spacji białych
'
'
      'Sub Findfirstcharacterinpara()
Dim wdoc As Document
Dim paral As Paragraph
Set wdoc = ActiveDocument
For Each paral In wdoc.Paragraphs
If paral.Range.Characters(1) = Chr(160) Then paral.Range.Characters(1).Delete
Next paral

    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^s^p"
        .Replacement.Text = "^p"
        .Forward = True
        .MatchWildcards = False
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
  
    
    
    'Sub podpisy_xml()

    
      With Selection.Find
    .Text = "Mikołaj Froń"
        .Replacement.Text = "^p<Author>^&</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
     
     With Selection.Find
    .Text = "Wojciech Naja"
        .Replacement.Text = "^p<Author>^&</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "Marcin Jastrzębski"
        .Replacement.Text = "^p<Author>Marcin Jastrzębski</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tso"
        .Replacement.Text = "^p<Author>Sławomir Oskarbski</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "^tdc"
        .Replacement.Text = "^p<Author>Dominika Czyż</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    With Selection.Find
        .Text = "^tol"
        .Replacement.Text = "^p<Author>Aleksandra Goles</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tmf"
        .Replacement.Text = "^p<Author>Mikołaj Froń</Author>^p"
        .Wrap = wdFindContinue
         End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
       
        .Text = "^tjas"
        .Replacement.Text = "^p<Author>Marcin Jastrzębski</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
     With Selection.Find
    .Text = "^ttab"
        .Replacement.Text = "^p<Author>Władysław Tabasz</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
     With Selection.Find
    .Text = "^ting"
        .Replacement.Text = "^p<Author>Inga Serafin</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
    .Text = "^tbo"
        .Replacement.Text = "^p<Author>Paweł Bochenek</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
    .Text = "^tnaj"
        .Replacement.Text = "^p<Author>Wojciech Naja</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
    .Text = "^tdc"
        .Replacement.Text = "^p<Author>Dominika Czyż</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      
     With Selection.Find
    .Text = "Ortalion"
        .Replacement.Text = "^p<Author>Bronisław Róg</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
         With Selection.Find
    .Text = "^tszy"
        .Replacement.Text = "^p<Author>Szymon Pacyna</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    'Sub lead_xml()
             
               
    Dim lRng As Range
    Set lRng = ActiveDocument.Range
    With lRng.Find
         
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphJustify
                    
                End With
                
        
        Do While .Execute(findText:=ChrW(9658), MatchWholeWord:=True)
                
                  
    
    lRng.Start = lRng.Paragraphs(1).Range.Start
     
     
            With lRng
                .InsertBefore "<Lead>"
            End With
            
         lRng.End = lRng.Paragraphs(1).Range.End - 1
             With lRng
                .InsertAfter "</Lead>"
                .Collapse 0
            End With
            
            
        Loop
    End With
lbll_Exit:
    Set lRng = Nothing
    'Exit Sub
    
 
 'Sub subtitle_loop_xml()

 
             
               
    Dim sRng As Range
    Set sRng = ActiveDocument.Range
    With sRng.Find
         
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphCenter
                End With
                 With .Font
                    .Size = 10
                    .Bold = True
                End With
  
        
        Do While .Execute(MatchWholeWord:=True)
                  
    
    sRng.Start = sRng.Paragraphs(1).Range.Start
     
     
            With sRng
                .InsertBefore "<Subtitle>"
            End With
            
         sRng.End = sRng.Paragraphs(1).Range.End - 1
             With sRng
                .InsertAfter "</Subtitle>"
                .Collapse 0
            End With
            
            
        Loop
    End With
lbls_Exit:
    Set sRng = Nothing
   'Exit Sub
 
    
    '<Body></Body>
   
            
   Dim bRng As Range
    Set bRng = ActiveDocument.Range
    With bRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphJustify
                End With
                 With .Font
                    .Size = 9
                    .Bold = False
                    .Italic = False
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    bRng.Start = bRng.Paragraphs(1).Range.Start
              With bRng
                .InsertBefore "<Body>"
            End With
            
         bRng.End = bRng.Paragraphs(1).Range.End - 1
             With bRng
                .InsertAfter "</Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lblb_Exit:
    Set bRng = Nothing
   'Exit Sub
   
   With Selection.Find
    .Text = "<Body><Author>"
        .Replacement.Text = "<Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
   With Selection.Find
    .Text = "</Author></Body>"
        .Replacement.Text = "</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   'Sub italic_xml()


Dim iPar As Paragraph
    For Each iPar In ActiveDocument.Range.Paragraphs
       
            If iPar.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft Or _
            iPar.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify Then
               If iPar.Range.Font.Size = 9 _
               And iPar.Range.Font.Name = "Georgia" _
               And iPar.Range.Font.Italic = True _
               And Len(iPar.Range) > 1 _
               And Not iPar.Range.Characters.First = "<" Then
                iPar.Range.InsertBefore "<Body originalStyle=""italic"">"
                iPar.Range.Characters.last.InsertBefore "</Body>"
              
            
            End If
            End If
       
    Next
    
    
    
    
    
    '<Description></Description>
   
   'Sub description_loop_xml()
          
    Dim dRng As Range
    Set dRng = ActiveDocument.Range
    With dRng.Find
        
                 'With .ParagraphFormat
                    '.Alignment = wdAlignParagraphLeft
                'End With
                 With .Font
                    .Size = 8
                    .Italic = True
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    dRng.Start = dRng.Paragraphs(1).Range.Start
              With dRng
                .InsertBefore "<Description>"
            End With
            
         dRng.End = dRng.Paragraphs(1).Range.End - 1
             With dRng
                .InsertAfter "</Description></Picture>"
                .Collapse 0
            End With
         
        Loop
    End With
lbld_Exit:
    Set dRng = Nothing
   'Exit Sub
   
   
    
    
    'Sub titlte_xml()
Application.ScreenUpdating = False
Dim i As Long
With ActiveDocument.Range
  With .Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Forward = True
    .Format = True
    .Text = "(*)^13"
    .Replacement.Text = "</Story>^p</Article>^p<Article>^p<Story>^p<Title>\1</Title>^p"
    .MatchWildcards = True
    .Wrap = wdFindContinue
    For i = 26 To 144
      .Font.Size = i / 2
      .Execute Replace:=wdReplaceAll
    Next
  End With
  With ActiveDocument.Range
    With .Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Forward = True
      .Format = True
      .Text = "<Title></Title>"
      .Replacement.Text = ""
      .Wrap = wdFindContinue
      .Execute Replace:=wdReplaceAll
    End With
  End With
End With
Application.ScreenUpdating = True






'Sub vignette_sport_xml()
             
    Dim vRng As Range
    Set vRng = ActiveDocument.Range
    With vRng.Find
         
        Do While .Execute(findText:=ChrW(9658), MatchWholeWord:=True)
                With .ParagraphFormat
                    .Alignment = wdAlignParagraphCenter
                End With
                'With .Font
                '    .Name = "Times New Roman"
               '  End With
    vRng.Start = vRng.Paragraphs(1).Range.Start
     
            With vRng
                                               
                .InsertBefore "<Vignette>"
            End With
            
         vRng.End = vRng.Paragraphs(1).Range.End - 1
             With vRng
                .InsertAfter "</Vignette>"
                .Collapse 0
            End With
                  
        Loop
    End With
lblf_Exit:
    Set vRng = Nothing
    'Exit Sub
    
    
    'Sub vignette_xml()
             
    Dim gRng As Range
    Set gRng = ActiveDocument.Range
    With gRng.Find
         
        Do While .Execute(findText:=ChrW(9660), MatchWholeWord:=True)
     
    gRng.Start = gRng.Paragraphs(1).Range.Start
     
            With gRng
                .InsertAfter "</Vignette>"
                .InsertBefore "<Vignette>"
                
                With .ParagraphFormat
                    
                End With
                .Collapse 0
            End With
                  
        Loop
    End With
lblg_Exit:
    Set gRng = Nothing
    'Exit Sub

'Sub title_remove_empty_paragraph_xml()

    

  
  
  'Sub FOTO_xml_loop_filename()
             
    
             
    Dim jjRng  As Range
    
    Set jjRng = ActiveDocument.Range
    
        Do While jjRng.Find.Execute(findText:=".JPG", _
                          MatchWholeWord:=True, MatchCase:=False)


            Set jjRng = jjRng.Paragraphs.Item(1).Range
            
            'fRng.Select 'tylko na czas testów!!!
            
            With jjRng
                .Text = "<Picture>" & Chr(13) & "<Image href=""file://images/" & Left(.Text, Len(.Text) - 1) & """></Image>" & String(2, Chr(13))
                .Collapse 0
            End With


        Loop


lbljj_Exit:
    Set jjRng = Nothing
    
        

'Sub FOTO_xml()
             
    Dim fRng As Range
    Set fRng = ActiveDocument.Range
    With fRng.Find
         
        Do While .Execute(findText:="FOTO", MatchWholeWord:=True)
     
    fRng.End = fRng.Paragraphs(1).Range.End
     
            With fRng
                .Font.Name = "Times New Roman"
                .Font.Size = 5
                '.Font.Italic = wdToggle
                .Bold = False
                .Font.SmallCaps = False
                .Font.AllCaps = False
                .InsertAfter "</Credits>" & Chr(13) & "</Picture>" & Chr(13)
                .InsertBefore "<Credits>"
                                  
              
                .Collapse 0
            End With
            
    
            'Exit Do
        Loop
    End With
    
lblj_Exit:
    Set fRng = Nothing
    'Exit Sub
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Picture>^p<Description>"
        .Replacement.Text = "<Description>"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


  Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Subtitle>^p<Subtitle>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'End Sub
    
    
    ' usuwanie dubli
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Body><Lead>"
        .Replacement.Text = "<Body>"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Lead></Body>"
        .Replacement.Text = "</Body>"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Lead>^p<Lead>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Subtitle>^p<Subtitle>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette><Lead>"
        .Replacement.Text = "<Lead>"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Lead></Vignette>"
        .Replacement.Text = "</Lead>"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    
    
    
    
    'Sub tabela_komplet()


'Sub add1_before_xml()
Dim aRng As Range
Dim aPar As Paragraph
  Set aRng = ActiveDocument.Range
  With aRng.Find
    With .Font
      .Size = 8
      .Underline = wdUnderlineSingle
    End With
    Do While .Execute(MatchWholeWord:=True)
      aRng.End = aRng.Paragraphs(1).Range.End - 1
      With aRng
        .Select
        .InsertBefore "<tr><td>"
        .InsertAfter "</td></tr>"
        .Collapse 0
      End With
    Loop
  End With
  For Each aPar In ActiveDocument.Range.Paragraphs
    If aPar.Range.Characters.First = "-" And Not InStr(aPar.Range.Text, "<tr>") > 0 _
    And aPar.Range.Font.Size = 8 Then
      aPar.Range.InsertBefore "<tr><td>"
      aPar.Range.InsertAfter "</td></tr>"
    End If
  Next
lbla_Exit:
  Set aRng = Nothing
  
  'tabela_tabulator
  
   Dim xRng As Range
    Set xRng = ActiveDocument.Range
    With xRng.Find
        
                
                 With .Font
                    .Size = 8
                    '.bold = False
                    '.Italic = False
                    .Underline = wdUnderlineSingle
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                         
             With xRng.Find
                .Text = "^t"
                .Font.Name = "Arial Narrow"
                .Replacement.Text = "</td><td>"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
         
        Loop
    End With
lblh_Exit:
    Set xRng = Nothing


'Sub dodaj_znacznik_tabeli_xml()
Dim zRng As Range
'Dim zPar As Paragraph
  Set zRng = ActiveDocument.Range
  With zRng.Find
      With .Font
      .Size = 8
      .Underline = wdUnderlineSingle
       End With
       With zRng.Find
       .Text = "<tr><td>1."
       End With
 
    
    Do While .Execute(MatchWholeWord:=True)
      zRng.End = zRng.Paragraphs(1).Range.End - 1
      With zRng
        .Select
        .InsertBefore "<Body><![CDATA[<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:270px;""><tbody>"
        .Collapse 0
      End With
    Loop
  End With
  
lblz_Exit:
  Set zRng = Nothing


'Sub add_sign_xml22()
Dim sPar As Paragraph
    For Each sPar In ActiveDocument.Range.Paragraphs
        If Not sPar.Range = ActiveDocument.Range.Paragraphs.last.Range Then
            If sPar.Range.Characters.First = "<" _
               And sPar.Range.Underline = wdUnderlineSingle _
               And Not InStrRev(sPar.Range.Text, "-") = 1 _
               And Not InStr(sPar.Next(1).Range.Text, "<tr>") > 0 _
               And Not sPar.Next(1).Range.Underline = wdUnderlineSingle _
               And sPar.Range.Font.Size = 8 Then
                
                
                sPar.Range.Characters.last.InsertBefore "</tbody></table>]]></Body>"
            End If
        End If
    Next
    
    
    'Sub body_arial()
   
            
    Dim aaRng As Range
    Set aaRng = ActiveDocument.Range
    With aaRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphJustify
                End With
                 With .Font
                    .Size = 8
                    '.bold = True
                    .Name = "Arial Narrow"
                    .Underline = wdUnderlineNone
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    aaRng.Start = aaRng.Paragraphs(1).Range.Start
              With aaRng
                .InsertBefore "<Body originalStyle=""body_ramka"">"
            End With
            
         aaRng.End = aaRng.Paragraphs(1).Range.End - 1
             With aaRng
                .InsertAfter "</Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lblbab_Exit:
    Set aaRng = Nothing
    
    
    
    
   'Sub bodybodybold()

   
Dim lPar As Paragraph
    For Each lPar In ActiveDocument.Range.Paragraphs
       
            If lPar.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify _
               And lPar.Range.Font.Bold = True _
               And lPar.Range.Font.Size = 9 _
               And Len(lPar.Range) > 2 _
               And Not lPar.Range.Characters.First = "<" Then
                lPar.Range.InsertBefore "<Body originalStyle=""bold"">"
                lPar.Range.Characters.last.InsertBefore "</Body>"
              
            
               
            End If
       
    Next
    
    
    Dim jPar As Paragraph
    For Each jPar In ActiveDocument.Range.Paragraphs
       
            If jPar.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft _
               And jPar.Range.Font.Bold = True _
               And jPar.Range.Font.Size = 9 _
               And Len(jPar.Range) > 2 _
               And Not jPar.Range.Characters.First = "<" Then
                jPar.Range.InsertBefore "<Body originalStyle=""bold"">"
                jPar.Range.Characters.last.InsertBefore "</Body>"
              
            
               
            End If
       
    Next
    
    
    'Sub arial_center()

   
Dim cPar As Paragraph
    For Each cPar In ActiveDocument.Range.Paragraphs
       
            If cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter _
              And Len(cPar.Range) > 2 _
               And cPar.Range.Font.Name = "Arial Narrow" Then
                cPar.Range.InsertBefore "<Body originalStyle=""body_ramkaj"">"
                cPar.Range.Characters.last.InsertBefore "</Body>"
              
            
               
            End If
       
    Next
    
    'Sub horoskop_xml()
Dim hPar As Paragraph
Dim hRng As Range
    For Each hPar In ActiveDocument.Range.Paragraphs
        Set hRng = hPar.Range
        hRng.End = hRng.End - 1
        If hRng.Font.Bold = True _
           And Len(hRng) > 1 _
           And hRng.Font.Color = RGB(0, 109, 53) Then
            hRng.InsertAfter "</Subtitle>"
            hRng.InsertBefore "<Subtitle originalStyle=""horoskop_znak"">"
        End If
    Next hPar
    Set hPar = Nothing
    Set hRng = Nothing
    
    'Sub bibliografia_xml()
Dim iiPar As Paragraph
Dim iiRng As Range
    For Each iiPar In ActiveDocument.Range.Paragraphs
        Set iiRng = iiPar.Range
        iiRng.End = iiRng.End - 1
        If iiRng.Font.Size = 9 _
           And iiRng.Font.Italic = True _
           And iiRng.ParagraphFormat.Alignment = wdAlignParagraphRight _
           And Len(iiRng) > 1 _
            Then
            iiRng.InsertAfter "</Body>"
            iiRng.InsertBefore "<Body originalStyle=""bibliografia"">"
        End If
    Next iiPar
    Set iiPar = Nothing
    Set iiRng = Nothing
    
    'Sub dowcipy_xml()
Dim dPar As Paragraph
Dim doRng As Range
    For Each dPar In ActiveDocument.Range.Paragraphs
        Set doRng = dPar.Range
        doRng.End = doRng.End - 1
        If doRng.Font.Size = 9 _
           And doRng.Font.Italic = True _
           And doRng.ParagraphFormat.Alignment = wdAlignParagraphCenter _
           And Len(doRng) > 1 _
            Then
            doRng.InsertAfter "</Body>"
            doRng.InsertBefore "<Body originalStyle=""dowcipy"">"
        End If
    Next dPar
    Set dPar = Nothing
    Set doRng = Nothing
    
    
    'Sub nazwy_miejscowosci_kolor_czerwony()


      'Application.ScreenUpdating = False
    Dim x As Long, ii As Long, ArrFnd()
    ArrFnd = Array("ROPCZYCE", "IWIERZYCE", "OSTRÓW", "BĘDZIENICA", "BYSTRZYCA", "NOCKOWA", _
    "OLCHOWA", "OLIMPÓW", "SIELEC", "WIERCANY", "WIŚNIOWA", "BRZEZÓWKA", "GNOJNICA", "LUBZINA", _
    "MAŁA", "NIEDŹWIADA", "OKONIN", "BLIZNA", "KAMIONKA", "KOZODRZA", "OCIEKA", "SKRZYSZÓW", _
    "ZDŻARY", "BĘDZIEMYŚL", "BORECZEK", "BUKOWINA", "CIERPISZ", "KAWĘCZYN", _
     "KLĘCZANY", "KRZYWA", "RUDA", "SZKODNA", "TARNÓW", _
    "ZABŁOCIE", "ZAGORZYCE", "BRONISZÓW", "BRZEZINY", "GLINIK", "NAWSIE", "RZESZÓW", "PASZCZYNA", _
    "WIELOPOLE SKRZYŃSKIE", "SĘDZISZÓW MŁP.", "BOREK WIELKI", "BOREK MAŁY", "WARSZAWA", _
"ŁĄCZKI KUCHARSKIE", "CZARNA SĘDZISZOWSKA", "WOLICA ŁUGOWA", "WOLICA PIASKOWA", _
"GÓRA ROPCZYCKA", "KAWĘCZYN SĘDZISZOWSKI", "WOLA OCIECKA", "CAŁY POWIAT", "SĘDZISZÓW MAŁOPOLSKI", _
"SIATKÓWKA", "SUMO", "TENIS STOŁOWY", "BOKS", "SZACHY", "KARATE", "HALOWA PIŁKA NOŻNA", _
    "PIŁKA NOŻNA", "PODNOSZENIE CIĘŻARÓW", "PIŁKARSKIE WIEŚCI", "ZAPASY", "KOLARSTWO")
    For x = 0 To UBound(ArrFnd)
        With ActiveDocument.Range
            With .Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ArrFnd(x)
                .Highlight = False
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = True
                .Font.Bold = True
                .Font.Name = "Georgia"
                .Execute
            End With
            Do While .Find.Found
                'i = i + 1
                '.Start = .Words.First.Start
                '.End = .Words.First.End
                '.MoveEndWhile " ", -1
                
                 .End = .End + 2
                .InsertAfter "{/Body:red}"
                .InsertBefore "{Body:red}"
                ' .End = .End + 1
                '.Font.Color = 204
                '.Font.bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
    'Application.ScreenUpdating = True
    'MsgBox i & " instances found."
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\<Lead\>\{Body:red\}"
        .Replacement.Text = "<Lead>{Lead:lead_red}"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\{Lead:lead_red\}*\{\/Body:red\}"
        .Replacement.Text = "^&{/Lead:lead_red}"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\{\/Body:red\}\{\/Lead:lead_red\}"
        .Replacement.Text = "{/Lead:lead_red}"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Image href=""file://images/<Body>"
        .Replacement.Text = "<Image href=""file://images/"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "file://images/<Body>"
        .Replacement.Text = "file://images/"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Body>""></Image>"
        .Replacement.Text = """></Image>"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Body originalStyle=""body_ramka""><Picture></Body>"
        .Replacement.Text = "<Picture>"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Image></Body>"
        .Replacement.Text = "</Image>"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Body originalStyle=""body_ramka""><Image"
        .Replacement.Text = "<Image"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Description></Picture>""></Image>"
        .Replacement.Text = "</Image>"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Image href=""file://images/<Description>"
        .Replacement.Text = "<Image href=""file://images/"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'Sub przepis_xml()

    
Dim ePar As Paragraph
Dim eRng As Range
    For Each ePar In ActiveDocument.Range.Paragraphs
        Set eRng = ePar.Range
        eRng.End = eRng.End - 1
        If Len(eRng) > 10 _
           And eRng.Font.Name = "Georgia" _
           And eRng.Font.Size = 10 _
            And Not eRng.Font.Bold = True _
           And eRng.ParagraphFormat.Alignment = wdAlignParagraphCenter Then
            'sPar.Range.HighlightColorIndex = wdDarkRed
            eRng.InsertAfter "</Subtitle>"
           eRng.InsertBefore "<Subtitle>"
        End If
    Next ePar
    Set ePar = Nothing
    Set eRng = Nothing
    
    
    

    
  'Sub add_xml_intro_outro1()


ActiveDocument.Content.InsertBefore "<?xml version='1.0' encoding='UTF-8' standalone='no'?>" & Chr(13) & "<Root>" & Chr(13)
ActiveDocument.Content.InsertAfter "</Story>" & Chr(13) & "</Article>" & Chr(13) & "</Root>"

'usuwanie <Root>^p</Story>^p</Article>
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Root>^p</Story>^p</Article>"
        .Replacement.Text = "<Root>"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    'usuwanie </Title>^p</Story>^p</Article>^p<Article>^p<Story>^p<Title>
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Title>^p</Story>^p</Article>^p<Article>^p<Story>^p<Title>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'horoskop dowcipy
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>DOWCIPY" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "</Story></Article><Article><Story><Title>Dowcipy</Title>"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>HOROSKOP REPORTERA" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "</Story></Article><Article><Story><Title>Horoskop Reportera</Title>"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>OPINIE" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "<Subtitle originalStyle=""oglo_tyt"">OPINIE" & ChrW(9660) & "</Subtitle>"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>OPINIA" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "<Subtitle originalStyle=""oglo_tyt"">OPINIA" & ChrW(9660) & "</Subtitle>"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>ZDANIEM ZAWODNIKA" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "<Subtitle originalStyle=""oglo_tyt"">ZDANIEM ZAWODNIKA" & ChrW(9660) & "</Subtitle>"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>ZDANIEM TRENERA" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "<Subtitle originalStyle=""oglo_tyt"">ZDANIEM TRENERA" & ChrW(9660) & "</Subtitle>"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>ZDANIEM PREZESA" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "<Subtitle originalStyle=""oglo_tyt"">ZDANIEM PREZESA" & ChrW(9660) & "</Subtitle>"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.WholeStory
    Selection.Fields.Unlink
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<tr><td>m.jastrzebski@reportergazeta.pl</td></tr>"
        .Replacement.Text = "mail"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Prosimy traktować poniższy horoskop z przymrużeniem oka, gdyż prawdopodobieństwo, że opisane sytuacje kiedykolwiek zaistnieją jest znikome i zależy jedynie od przypadku."
        .Replacement.Text = "<Body originalStyle=""dowcipy"">^&</Body><Body><![CDATA[<br />]]></Body>"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Wyrażane przez Czytelników opinie nie są stanowiskiem redakcji Reporter Gazety."
        .Replacement.Text = "<Body originalStyle=""dowcipy"">^&</Body><Body><![CDATA[<br />]]></Body>"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'Sub zdjecie_oglo_xml_opinia_rrr11()



Dim parao As Paragraph
For Each parao In ActiveDocument.Paragraphs
    
    If InStr(parao, "OPINIA") > 0 Or InStr(parao, "OPINIE") > 0 _
    Or InStr(parao, "ZDANIEM ZAWODNIKA") > 0 _
    Or InStr(parao, "ZDANIEM TRENERA") > 0 _
    Or InStr(parao, "ZDANIEM PREZESA") > 0 Then
        
              If Len(parao.Range) > 2 _
               And parao.Range.Font.Name = "Times New Roman" _
               And parao.Range.Font.Italic = False Then
                parao.Next(3).Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)
              End If
                
    End If
Next

'Sub zdjecie_oglo_xml_opinia_za_kazdym_razem()



Dim parax As Paragraph
For Each parax In ActiveDocument.Paragraphs
    
     If Not parax.Range = ActiveDocument.Range.Paragraphs.First.Range _
     And Not parax.Range = ActiveDocument.Range.Paragraphs.last.Range Then
    If parax.Range.Next.Font.Name = "Georgia" _
               And parax.Range.Previous.Font.Name = "Georgia" _
               And parax.Range.Next.Font.Size = 9 _
               And parax.Range.Previous.Font.Size = 9 _
               And parax.Range.Previous.Font.Bold = True _
               And parax.Range.Next.Font.Italic = True _
                 And parax.Range.Next.ParagraphFormat.Alignment = wdAlignParagraphJustify _
                And parax.Range.Previous.ParagraphFormat.Alignment = wdAlignParagraphJustify _
                Then
                parax.Next(1).Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)
              End If
                
    End If
Next


'<Body></Body> italic first - normal last
   
               
Dim sPara As Paragraph
Dim sRnga As Range
    For Each sPara In ActiveDocument.Range.Paragraphs
        Set sRnga = sPara.Range
        sRnga.End = sRnga.End - 1
        If Len(sRnga) > 1 _
            And sRnga.Font.Size = 9 _
            And sRnga.Font.Bold = False _
           And sRnga.ParagraphFormat.Alignment = wdAlignParagraphJustify _
            And sRnga.Characters.First.Italic = True _
           And sRnga.Characters.last.Italic = False _
           And Not sRnga.Characters.First = "<" _
           And Not sRnga.Characters.last = ">" Then
            
            sRnga.InsertAfter "</Body>"
           sRnga.InsertBefore "<Body>"
        End If
    Next sPara
    Set sPara = Nothing
    Set sRnga = Nothing
    
    

' autorzy_niepodpisani
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>REPORTER I PIENIĄDZE" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "^&^p<Author>Paweł Bochenek</Author> "
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>REPORTER I PRAWO" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "^&^p<Author>Inga Serafin</Author> "
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>(TROCHĘ) MŁODSZYM OKIEM" & ChrW(9658) & " Mikołaj Froń </Vignette>"
        .Replacement.Text = "^&^p<Author>Mikołaj Froń</Author> "
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>CAŁKIEM (NIE) OBIEKTYWNIE" & ChrW(9658) & " Wojciech Naja </Vignette>"
        .Replacement.Text = "^&^p<Author>Wojciech Naja</Author> "
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>HISTORIA" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "^&^p<Author>Szymon Pacyna</Author> "
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "&"
        .Replacement.Text = "&amp;"
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    


'Sub highlight_xml()


Dim hiPar As Paragraph
    For Each hiPar In ActiveDocument.Range.Paragraphs
       
            If Not hiPar.Range.Characters.First = "<" _
            And Not hiPar.Range.Characters.last = ">" _
            Then
                hiPar.Range.HighlightColorIndex = wdBrightGreen
              
            
                       End If
       
    Next



End Sub
