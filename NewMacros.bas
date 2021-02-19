Attribute VB_Name = "NewMacros"

Sub AutoExec()
'
' AutoExec Makro
'
'

End Sub
Sub polsat()
'
' polsat Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", odc."
        .Replacement.Text = " ("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = ") "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "Dla ma³oletnich od lat 12) Udogodnienia: napisy dla nies³ysz¹cych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 12"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "Udogodnienia: napisy dla nies³ysz¹cych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 16"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 16"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 7"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 2
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "(^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "(^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(^#^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub TVN()
'
' TVN czyszczenie Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 16)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", live"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", Dolby"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", napisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^#^#^#^#) - informacje"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub tvp()
'
' tvp Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "; Bez ograniczeñ wiekowych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = _
            "(godz. 06:15, 06:45, 07:15); magazyn; STEREO, 16:9, Bez ograniczeñ wiekowych, Na ¿ywo"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "; STEREO, 16:9, Na ¿ywo"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "; STEREO, 16:9, Dla ma³oletnich od lat 12, Na ¿ywo"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "; Dla ma³oletnich od lat 12"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "; Dla ma³oletnich od lat 7"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " txt. str. 777; "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "kraj prod."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "; STEREO, 16:9"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "-  txt. str. 777"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", Dla ma³oletnich od lat 12"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Weekendowy Hit Jedynki -"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "STEREO, 16:9"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "; Dla ma³oletnich od lat 16"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     With Selection.Find
        .Text = ", Bez ograniczeñ wiekowych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", Na ¿ywo"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
'dodanie_nawiasow_tvp
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- odc. "
        .Replacement.Text = "("
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "odc. "
        .Replacement.Text = "("
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     With Selection.Find
        .Text = "- odc "
        .Replacement.Text = "("
        .Forward = True
        .Wrap = wdFindAsk
            
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ";"
        .Replacement.Text = ") - "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^#^#^#^#) - teleturniej muzyczny"
        .Replacement.Text = "- teleturniej"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Pogoda - "
        .Replacement.Text = "Pogoda"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ALARM!)"
        .Replacement.Text = "ALARM!"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Jeden z dziesiêciu - ^#/^#^#^#) - teleturniej"
        .Replacement.Text = "Jeden z dziesiêciu - teleturniej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Ko³o fortuny (^#^#^# ed. ^#) - teleturniej"
        .Replacement.Text = "Ko³o fortuny - teleturniej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Familiada (^#^#^#^#) - teleturniej"
        .Replacement.Text = "Familiada - teleturniej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Jeden z dziesiêciu - ^#^#/^#^#^#) - teleturniej"
        .Replacement.Text = "Jeden z dziesiêciu - teleturniej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Elif s.II"
        .Replacement.Text = "Elif"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Wszystko dla pañ s.II"
        .Replacement.Text = "Wszystko dla pañ"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Agropogoda)"
        .Replacement.Text = "Agropogoda"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Klan (^#^#^#^#"
        .Replacement.Text = "^&)"
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    ' usuwanie_podwojnych_nawiasow
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "))"
        .Replacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     ' usuwanie_podwojnych_spacji
      
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'rez_wyk_tvp
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ") - re¿.:"
        .Replacement.Text = "); re¿.: "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ") - wyk.:"
        .Replacement.Text = "; wyk.: "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Jaka to melodia? (^#^#^#^#) - teleturniej muzyczny"
        .Replacement.Text = "Jaka to melodia? - teleturniej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^#^# "
        .Replacement.Text = "^&)"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^#/^#^# " '(6/13
        .Replacement.Text = "^&)"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " )"
        .Replacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- (N) - "
        .Replacement.Text = "- "
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(N)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " - telenowela historyczna TVP"
        .Replacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ElifI"
        .Replacement.Text = "Elif"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "s.IV "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "s.II "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "s.I "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "s.XIII "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "s.XXII "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Doktor z alpejskiej wioski - nowy rozdzia³"
        .Replacement.Text = "Doktor z alpejskiej..."
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
End Sub
Sub tv4()
'
' tv4 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 16"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Udogodnienia: napisy dla nies³ysz¹cych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 7"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Udogodnienia: audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " Dla ma³oletnich od lat 12"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", odc."
        .Replacement.Text = " ("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", odc."
        .Replacement.Text = " ("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
        .Text = "^l"
        .Replacement.Text = ") "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Bez ograniczeñ wiekowych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ") )"
        .Replacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub tv4niepogrub()
'
' tv4niepogrub Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = False
        .Italic = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = False
        .Italic = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
End Sub
Sub tvn7()
'
' tvn7 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- program obyczajowy (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial , USA (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial inne (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- program (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- serial obyczajowy, Polska (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 12), napisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial , USA (dla ma³oletnich od lat 16), napisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- program (dla ma³oletnich od lat 16)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- talk show (dla ma³oletnich od lat 12), audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", napisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- program s¹dowy (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial S-F, USA (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = _
            "- program kryminalny (dla ma³oletnich od lat 12), audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial , USA (dla ma³oletnich od lat 16)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 16)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 12), Dolby"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", Dolby"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub formatarial()
'
' formatarial Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 7.5
        .Kerning = 0
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 7.5
        .Kerning = 0
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 7.5
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 7.5
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 7.5
    WordBasic.OpenOrCloseParaBelow
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub AddBrackets()
'
' AddBrackets Makro
'
'

End Sub
Sub nawiasy_polsat()
'
' nawiasy_polsat Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = False
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = ")"
        .Replacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "(^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "(^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "(^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "(^#^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
Sub tvn_pogrubienia()
'
' tvn_pogrubienia
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Uwaga!"
        .Replacement.Text = "Uwaga!"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Mango - Telezakupy"
        .Replacement.Text = "Mango - Telezakupy"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Kuchenne rewolucje 11"
        .Replacement.Text = "Kuchenne rewolucje 11"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Doradca smaku"
        .Replacement.Text = "Doradca smaku"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Dzieñ Dobry TVN"
        .Replacement.Text = "Dzieñ Dobry TVN"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Ukryta prawda"
        .Replacement.Text = "Ukryta prawda"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Szko³a"
        .Replacement.Text = "Szko³a"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "19 +"
        .Replacement.Text = "19 +"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Szpital"
        .Replacement.Text = "Szpital"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Kuchenne rewolucje 4"
        .Replacement.Text = "Kuchenne rewolucje 4"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Szko³a"
        .Replacement.Text = "Szko³a"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Fakty"
        .Replacement.Text = "Fakty"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Sport"
        .Replacement.Text = "Sport"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Pogoda"
        .Replacement.Text = "Pogoda"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Kuba Wojewódzki 12"
        .Replacement.Text = "Kuba Wojewódzki 12"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Revolution II"
        .Replacement.Text = "Revolution II"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Moc Magii"
        .Replacement.Text = "Moc Magii"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Nic straconego"
        .Replacement.Text = "Nic straconego"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Na Wspólnej Omnibus 15"
        .Replacement.Text = "Na Wspólnej Omnibus 15"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Drzewo marzeñ"
        .Replacement.Text = "Drzewo marzeñ"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "MasterChef 6"
        .Replacement.Text = "MasterChef 6"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Azja Express 2"
        .Replacement.Text = "Azja Express 2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "36,6 2"
        .Replacement.Text = "36,6 2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Mam talent 10"
        .Replacement.Text = "Mam talent 10"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Nowa Maja w ogrodzi"
        .Replacement.Text = "Nowa Maja w ogrodzi"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Nowa Maja w ogrodzie"
        .Replacement.Text = "Nowa Maja w ogrodzie"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Akademia ogrodnika"
        .Replacement.Text = "Akademia ogrodnika"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Kobieta na krañcu œwiata 9"
        .Replacement.Text = "Kobieta na krañcu œwiata 9"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Co za tydzieñ"
        .Replacement.Text = "Co za tydzieñ"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Diagnoza ^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Druga szansa ^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Na Wspólnej 15"
        .Replacement.Text = "Na Wspólnej 15"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Milionerzy"
        .Replacement.Text = "Milionerzy"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "The Following"
        .Replacement.Text = "The Following"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Œlub od pierwszego wejrzenia 2"
        .Replacement.Text = "Œlub od pierwszego wejrzenia 2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Superwizjer"
        .Replacement.Text = "Superwizjer"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub godziny_pogrubienie()
'
' godziny_pogrubienie Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub pogrubienia_TVP()
'
' pogrubienia_TVP Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "TELEZAKUPY"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Jaka to melodia?"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Dzieñ dobry Polsko!"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "Korona królów"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "Program rozrywkowy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "Wiadomoœci"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Pogoda poranna"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Kwadrans polityczny"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Komisariat"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Zakochaj siê w Polsce"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Rodzinny ekspres"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Pó³noc - Po³udnie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Dr Quinn"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Agrobiznes"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Agropogoda"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Nowoczesnoœæ w rolnictwie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "£owcy. Na otwartej przestrzeni"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Elif"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Opole 2017 na bis"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Ktokolwiek widzia³..."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Pogoda"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "W sercu miasta"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Teleexpress"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Klan"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Rodzina wie lepiej"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Mam prawo"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Sport"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "The Wall. Wygraj marzenia"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Notacje"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Zakoñczenie dnia"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Galeria"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Sprawa dla reportera"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Naszaarmia.pl"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Rok w ogrodzie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Rok w ogrodzie extra"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Ptaki ciernistych krzewów"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Studio Raban"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Okrasa ³amie przepisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Ojciec Mateusz"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Zagadka Hotelu Grand"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Rolnik szuka ¿ony seria IV"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Dziewczyny ze Lwowa"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Komisarz Alex"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "S³ownik polsko@polski"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Pe³nosprawni"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Transmisja Mszy Œwiêtej"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Tydzieñ"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Ziarno"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Jak to dzia³a"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Biblia"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Weterynarze z sercem"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Sekrety mnichów"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Miêdzy ziemi¹ a niebem"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Anio³ Pañski"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Poldark - Wichry losu"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Sonda 2"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Obserwator"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "To siê op³aca"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Magazyn œledczy Anity Gargas"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Smaki polskie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Rodzinka.pl"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Sztuka codziennoœci"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "M jak mi³oœæ"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Pytanie na œniadanie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Panorama"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Barwy szczêœcia"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Postaw na milion"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Tylko z Tob¹"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Na sygnale"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Coœ dla Ciebie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Janosik"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Ko³o fortuny"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Nadzieja i mi³oœæ"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Sport Telegram"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Jeden z dziesiêciu"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "O mnie siê nie martw"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Pod wspólnym niebem"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Na sygnale"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Na sygnale"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Podró¿e z histori¹"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Familiada"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "By³o... nie minê³o"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Pierwsza randka"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Na dobre i na z³e"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "S³owo na niedzielê"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Kulisy - Postaw na milion"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "The Voice of Poland"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Miasto skarbów"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Grupa specjalna ""Kryzys"""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "la nies³ysz¹cych - S³owo na niedzielê"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 63
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Dla nies³ysz¹cych - S³owo na niedzielê"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 63
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Ukryte skarby"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 63
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Przeprowadzki zwierz¹t"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 64
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Mak³owicz w podró¿y"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 63
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Bake off - Ale ciacho!"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Bake off - Ale przepis"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 65
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Lajk!"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 70
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Zakoñczenie dnia"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 76
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Kulisy serialu ""M jak mi³oœæ"""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 73
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Dla nies³ysz¹cych - M jak mi³oœæ"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 77
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Magazyn Ekspresu Reporterów"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 76
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Zakoñczenie programu"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 82
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Licencja na wychowani"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Prokurator"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Licencja na wychowanie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Licencja na wychowanie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Downton Abbey"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Ktokolwiek widzia³, ktokolwiek wie"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
          Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Bia³o - czerwoni"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Skoki Narciarskie - Puchar Œwiata"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "ALARM!"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
       
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Magazyn Rolniczy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
       
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Leœniczówka"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
       
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' godziny_pogrubienie Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
End Sub
Sub tv4_do_test()
'
' tv4_do_test Makro
'

'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 16"
        .Replacement.Text = ""
       .Execute Replace:=wdReplaceAll

        .Text = "Udogodnienia: napisy dla nies³ysz¹cych"
        .Replacement.Text = ""
       .Execute Replace:=wdReplaceAll
  
        .Text = "Dla ma³oletnich od lat 7"
        .Replacement.Text = ""
       .Execute Replace:=wdReplaceAll
  .Execute Replace:=wdReplaceAll
  
        .Text = "Udogodnienia: audiodeskrypcja"
        .Replacement.Text = ""
       .Execute Replace:=wdReplaceAll
    
        .Text = " Dla ma³oletnich od lat 12"
        .Replacement.Text = ""
       .Execute Replace:=wdReplaceAll
    
        .Text = ", odc."
        .Replacement.Text = " ("
       .Execute Replace:=wdReplaceAll
  
        .Text = ", odc."
        .Replacement.Text = " ("
       .ClearFormatting

        .Text = "^l"
        .Replacement.Text = ") "
        .Execute Replace:=wdReplaceAll

        .Text = " "
        .Replacement.Text = " "
       .Execute Replace:=wdReplaceAll
    
        .Text = "Bez ograniczeñ wiekowych"
        .Replacement.Text = ""
       .Execute Replace:=wdReplaceAll
    
        .Text = ") )"
        .Replacement.Text = ")"
       .Execute Replace:=wdReplaceAll
End With
End Sub

Sub kompletny_tv4()
'
' kompletny tv4
'
'czyszczenie

Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 16"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Udogodnienia: napisy dla nies³ysz¹cych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 7"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Udogodnienia: audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " Dla ma³oletnich od lat 12"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", odc."
        .Replacement.Text = " ("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", odc."
        .Replacement.Text = " ("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
        .Text = "^l"
        .Replacement.Text = ") "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Bez ograniczeñ wiekowych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ") )"
        .Replacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "i audiodeskrypcja "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' godziny_pogrubienie
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


'niepogrubienie


 Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = False
        .Italic = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = False
        .Italic = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' wszystko bez blod

Selection.WholeStory
    Selection.Font.Bold = wdToggle
    Selection.MoveUp Unit:=wdLine, Count:=1

' godziny_pogrubienie
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' podwójne nawiasy ))

     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ") )"
        .Replacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "16:1521:10"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' powrot_na_poczatek_dokumentu
    
        Selection.WholeStory
    Selection.MoveUp Unit:=wdLine, Count:=1
    
    ' odtabelowanie_tvn
'
'
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=1
        
'   formatarial Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 7.5
        .Kerning = 0
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 7.5
        .Kerning = 0
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 7.5
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 7.5
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 7.5
    WordBasic.OpenOrCloseParaBelow
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "i audiodeskrypcja "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ") Dla ma³oletnich od lat 12"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


End Sub

Sub Makro_trwam()
'
' Makro24 Makro
'
'usuwanie linijek

    Selection.Rows.Delete
    Selection.Rows.Delete
    Selection.Rows.Delete
    Selection.Rows.Delete
    Selection.Rows.Delete
    Selection.Rows.Delete
    Selection.Rows.Delete
    Selection.Rows.Delete
    Selection.Rows.Delete
    Selection.Rows.Delete
     
    'usuwanie kolumn
    
    ActiveDocument.Tables(1).Columns(2).Select
    Selection.Columns.Delete
    ActiveDocument.Tables(1).Columns(2).Select
    Selection.Columns.Delete
    ActiveDocument.Tables(1).Columns(2).Select
    Selection.Columns.Delete
    ActiveDocument.Tables(1).Columns(3).Select
    Selection.Columns.Delete
    ActiveDocument.Tables(1).Columns(3).Select
    Selection.Columns.Delete
    ActiveDocument.Tables(1).Columns(3).Select
    Selection.Columns.Delete
    ActiveDocument.Tables(1).Columns(3).Select
    Selection.Columns.Delete
    Selection.Columns.Delete
 
    ' wszystko bez blod

Selection.WholeStory
    Selection.Font.Bold = wdToggle
    Selection.MoveUp Unit:=wdLine, Count:=1

' godziny_pogrubienie
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'pogrubiebie godziny w formacie 0:00
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'konwersja tabeli
    
    ActiveDocument.Tables(1).Columns(1).Select
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveUp Unit:=wdLine, Count:=1
    
       
End Sub
Sub konwersja_tabeli_na_tekst()
'
' Makro25 Makro
'
'
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveUp Unit:=wdLine, Count:=1
End Sub
Sub test_dodawania_nawiasow()
'
' test_dodawania_nawiasow Makro
'
'
'Shape(1) holds "ABC DEF GHI JKL MNO"

    Dim c As Range
    Dim lPos As Long
    Dim sTextToFind As String
    
    

   sTextToFind = " DEF "

    Let c = ActiveDocument.TextFrame.TextRange
    lPos = InStr(c, sTextToFind) 'Returns position 4 (the space between C & D).

    'Returns "ABC DEF my new text GHI JKL MNO"
    c.Text = Left(c, lPos + Len(sTextToFind) - 1) & "my new text " & Mid(c, lPos + Len(sTextToFind))

End Sub


Sub kompletny_TVN()
'
' kompletny_TVN
'
'
' TVN czyszczenie Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 16)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", live"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", Dolby"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", napisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^#^#^#^#) - informacje"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ","
        .Replacement.Text = ", "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "("
        .Replacement.Text = " ("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "dla niedos³ysz¹cych [N]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[N]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[AD]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " ,"
        .Replacement.Text = ","
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Raport smogowy – wiem czym oddycham"
        .Replacement.Text = "Raport smogowy"
        .Forward = True
        '.Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


     
 ' powrot_na_poczatek_dokumentu
    
        Selection.WholeStory
    Selection.MoveUp Unit:=wdLine, Count:=1
    
    ' odtabelowanie_tvn
'
'
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=1
    
' usuniecie_interlinii

    Selection.WholeStory
    WordBasic.OpenOrCloseParaBelow
    Selection.MoveUp Unit:=wdLine, Count:=1
    
    ' tvn_pogrubienia

'Sub test2()
    Dim r As Range
    
    Set r = ActiveDocument.Range
    
    r.Font.Bold = True

    With r.Find
        .MatchWildcards = True
        
        .Text = "[\(\-]*^13"
        .Replacement.Font.Bold = False
        .Execute Replace:=wdReplaceAll
    End With
    
    
    ' program_formatowanie_czcionka_wciecie Makro
'
'
    Selection.WholeStory
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 8
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .FirstLineIndent = CentimetersToPoints(-0.66)
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
    End With
    
End Sub
Sub kompletny_TVN7()

' czyszczenie_tvn7
'
'
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", audiodeskrypcja [AD]"
        .Replacement.Text = ""
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- program obyczajowy (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial , USA (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial inne (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- program (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- serial obyczajowy, Polska (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 12), napisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial , USA (dla ma³oletnich od lat 16), napisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- program (dla ma³oletnich od lat 16)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- talk show (dla ma³oletnich od lat 12), audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ", napisy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- program s¹dowy (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial S-F, USA (dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = _
            "- program kryminalny (dla ma³oletnich od lat 12), audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "- serial , USA (dla ma³oletnich od lat 16)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 16)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 12), Dolby"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(dla ma³oletnich od lat 12)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", Dolby"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ","
        .Replacement.Text = ", "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "("
        .Replacement.Text = " ("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[N] "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "dla niedos³ysz¹cych [N] "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", audiodeskrypcja [AD] "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
  
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- program obyczajowy dla niedos³ysz¹cych [N]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- program s¹dowy"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- serial , Polska"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "- serial sensacyjny, USA [N]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[AD]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[N]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' godziny_pogrubienie Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
 ' powrot_na_poczatek_dokumentu
    
        Selection.WholeStory
    Selection.MoveUp Unit:=wdLine, Count:=1
    
    ' odtabelowanie_tvn
'
'
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=5
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    
    ' formatarial Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 7.5
        .Kerning = 0
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 7.5
        .Kerning = 0
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 7.5
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 7.5
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 7.5
    WordBasic.OpenOrCloseParaBelow
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub

Sub kompletny_polsat_cz1()
'
' polsat czyszczenie nawiasy
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", odc."
        .Replacement.Text = " ("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = ") "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "Dla ma³oletnich od lat 12) Udogodnienia: napisy dla nies³ysz¹cych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 12"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "Udogodnienia: napisy dla nies³ysz¹cych"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 16"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 16"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "Dla ma³oletnich od lat 7"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 2
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "(^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "(^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "(^#^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "i audiodeskrypcja"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "Bez ograniczeñ wiekowych)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    

' godziny_pogrubienie Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'obcinanie numerów tajemnice losu i disco gramy
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Tajemnice losu (^#^#^#^#)"
        .Replacement.Text = "Tajemnice losu"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Disco Gramy (^#^#^#^#)"
        .Replacement.Text = "Disco Gramy"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
End Sub

Sub kompletny_polsat_cz2()

' polsat_usuniecie_niepogrubionych
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = False
        .Italic = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = False
        .Italic = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

' usuwanie_podwójnych_spacji
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
' nawiasy_polsat Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = False
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = ")"
        .Replacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "(^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "(^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "(^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .Text = "(^#^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
 ' powrot_na_poczatek_dokumentu
    
        Selection.WholeStory
    Selection.MoveUp Unit:=wdLine, Count:=1
    
    
    ' odtabelowanie_polsat
'
'
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=5
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    
    'Sub test2()
    Dim r As Range
    
    Set r = ActiveDocument.Range
    
    r.Font.Bold = True

    With r.Find
        .MatchWildcards = True
        
        .Text = "[\(\-]*^13"
        .Replacement.Font.Bold = False
        .Execute Replace:=wdReplaceAll
    End With
    
    
    ' program_formatowanie_czcionka_wciecie Makro
'
'
    Selection.WholeStory
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 8
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .FirstLineIndent = CentimetersToPoints(-0.66)
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
    End With
    

End Sub
Sub pozbywanie_tabeli()
'
' pozbywanie_tabeli Makro
'
'
'pozbywanie tabel

       Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=1
        
End Sub
Sub powrot_na_poczatek()
'
' powrot_na_poczatek Makro
'
'
  ' powrot_na_poczatek_dokumentu
    
        
    Selection.WholeStory
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.MoveDown Unit:=wdLine, Count:=3
    
End Sub


Sub tv4_pozbywanie_tabeli_format_arial()
'
' tv4_pozbywanie_tabeli_format_arial Makro
'
''pozbywanie tabel

       Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=6
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveDown Unit:=wdLine, Count:=1
        
'   formatarial Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 7.5
        .Kerning = 0
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 7.5
        .Kerning = 0
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 7.5
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 7.5
    With Selection.Find
        .Text = ""
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 7.5
    WordBasic.OpenOrCloseParaBelow
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll


End Sub
Sub tabela_usuwanie()
'
' tabela_usuwanie Makro
'
'
    Selection.WholeStory
End Sub
Sub pwrotnapoczatek()
'
' pwrotnapoczatek Makro
'
'
    Selection.WholeStory
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
End Sub
Sub test_dodawania_elemetow()
'
' test_dodawania_elemetow Makro
'
'Sub S1204A_InsertTags()
    Dim parEach As Paragraph
    Dim rngEach As Range
    Dim strTag As String
    
    ' Process all paragraphs in selection
    For Each parEach In Selection.Paragraphs
        Set rngEach = parEach.Range
        
        ' Choose appropriate tag
        If rngEach.ComputeStatistics(wdStatisticLines) = 1 Then
            strTag = "h1"
        ElseIf parEach.LeftIndent > 0 And parEach.RightIndent > 0 Then
            strTag = " DEF "
        Else
            strTag = "p"
        End If
        
        ' Insert opening tag
        rngEach.InsertBefore Text:="<" & strTag & ">"
        
        ' Move end point of range before end of paragraph
        rngEach.MoveEnd Unit:=wdCharacter, Count:=-1
        
        ' Insert closing tag
        rngEach.InsertAfter Text:="</" & strTag & ">"
    Next parEach
End Sub


Sub godziny_pogrubienie_dodanie()
'
' godziny_pogrubienie Makro
'
'
        With Selection.Find
        .Text = "^#^#:^#^#"
        .Repleacement.Text = ")"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Private Sub AddBracketscall()
Dim selectionLength As Long
 selectionLength = Selection.Characters.Count
  
 Selection.MoveRight Unit:=wdCharacter, Count:=1 'cleared the selection and moved to the right end of it, now move to the beginning and select the whole thing again, from left to right this time!


Selection.MoveLeft Unit:=wdCharacter, Count:=selectionLength


 Selection.MoveRight Unit:=wdCharacter, Count:=selectionLength, Extend:=wdExtend
 
 Dim iCount As Integer
 iCount = 1
 While Right(Selection.Text, 1) = " " Or _
 Right(Selection.Text, 1) = Chr(13)
 Selection.MoveLeft Unit:=wdCharacter, Count:=1, _
  Extend:=wdExtend
 iCount = iCount + 1
 Wend

 Selection.InsertAfter ")"
 Selection.InsertBefore "("
 Selection.MoveRight Unit:=wdCharacter, Count:=iCount
End Sub


Sub AppendToExistingOnLeft()
'
' AppendToExistingOnLeft Makro
'
'

End Sub
Sub paragrafbold()
'
' paragrafbold Makro
'
Dim p As Paragraph
Dim oRng As Range
Dim Thistext As String
Dim meetssomecondition As Boolean

    For Each p In ActiveDocument.Paragraphs
        meetssomecondition = False
        Thistext = p.Range.Text
        If InStr(1, Thistext, "^l") > 0 Then meetssomecondition = True
        If meetssomecondition = True Then
            Set oRng = p.Range    'set a range to the paragraph
            oRng.Collapse 1    'collapse the range to its start
            oRng.MoveEndUntil "^l"    'move the end of the range to the hyphen
            oRng.Font.Bold = True    'format the range
        End If
    Next p
    End Sub

Sub dodawanie_znakow_demo()
'
' dodawanie_znakow_demo Makro
'
Application.ScreenUpdating = False
Dim i As Long
With ActiveDocument.Range
  With .Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = ""
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = True
    .Font.Italic = True
    .Execute
  End With
  Do While .Find.Found
    i = i + 1
    If .Characters.First = " " Then
      .Characters.First.Font.Italic = False
      .Start = .Start + 1
    End If
    .Characters.First.Previous.InsertBefore "§"
    If .Characters.last = " " Then
      .Characters.last.Font.Italic = False
      .End = .End - 1
    End If
    .Characters.last.InsertAfter "@"
    .Characters.last.Font.Italic = False
    .End = .End + 1
    'The next line is only needed if the Find is based on formatting without regard to text
    If .End = ActiveDocument.Range.End Then Exit Sub
    .Collapse wdCollapseEnd
    .Find.Execute
  Loop
End With
Application.ScreenUpdating = True
MsgBox i & " instances found."
End Sub

Sub dodawanie_znakow_demo_2()
'
' dodawanie_znakow_demo_2 Makro
'

Application.ScreenUpdating = False
Dim Tbl As Table, FndList, RepList, i As Long
FndList = Array("^#", "^#^#")
RepList = Array("(^#)", "(^#^#)")
With ActiveDocument
  For Each Tbl In .Tables
    With Tbl.Range.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .MatchWholeWord = True
      .Replacement.Text = "^&"
      .Wrap = wdFindStop
      For i = 0 To UBound(FndList)
        .Text = FndList(i)
        .Replacement.Text = RepList(i)
        .Execute Replace:=wdReplaceAll
      Next
    End With
  Next
End With
Application.ScreenUpdating = True
End Sub

Sub reporter()
'
'formatowanie czcionki i tabulatora

    Selection.WholeStory
    Selection.Font.Name = "Georgia"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.93) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    


'Graham Mayor - http://www.gmayor.com - Last updated - 12 Jan 2018
Dim oRng As Range
Dim strEndWord() As Variant
Dim i As Long
    strEndWord = Array("so", "jas", "mf")
    Set oRng = Selection.Paragraphs(1).Range
    With oRng
        .Font.Bold = True
        .End = .Words(1).End - 1
        .InsertAfter ChrW(9658)
        .End = .End + 1
        .Font.Color = 204
        .End = ActiveDocument.Range.End - 1
        .Start = .Words.last.Start
        For i = 0 To UBound(strEndWord)
            If strEndWord(i) = LCase(.Text) Then
                .InsertBefore vbTab
                .Font.Bold = True
                Exit For
            End If
        Next i
    End With
lbl_Exit:
    Set oRng = Nothing
    Exit Sub
End Sub

Sub rozmowa()
'
' rozmowa
'
'
    Selection.WholeStory
    Selection.Font.Name = "Georgia"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.93) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
End Sub
Sub test_reporter()
'
' test_reporter Makro
'
'
'reporter Makro
'
'Graham Mayor - http://www.gmayor.com - Last updated - 12 Jan 2018
Dim oRng As Range
Dim strEndWord() As Variant
Dim listam() As Variant
Dim j As Long
    strEndWord = Array("so", "jas", "mf")
    listam = Array("iwierzyce", "ropczyce")
          
     For j = 0 To UBound(listam)
            If listam(j) = UCase(.Text) Then
                .InsertAfter ChrW(9658)
                .Font.Bold = True
                Exit For
            End If
        Next j
    End With
lbl_Exit:
    Set oRng = Nothing
        
       
End Sub
Sub array_test()
'

'formatowanie czcionki i tabulatora

    Selection.WholeStory
    Selection.Font.Name = "Georgia"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.93) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces

'pogrubianie nazw miejscowosci
      
      'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("ROPCZYCE", "IWIERZYCE", "OSTRÓW", "BÊDZIENICA", "BYSTRZYCA", "NOCKOWA", _
    "OLCHOWA", "OLIMPÓW", "SIELEC", "WIERCANY", "WIŒNIOWA", "BRZEZÓWKA", "GNOJNICA", "LUBZINA", _
    "MA£A", "NIEDWIADA", "OKONIN", "BLIZNA", "KAMIONKA", "KOZODRZA", "OCIEKA", "SKRZYSZÓW", _
    "ZD¯ARY", "BÊDZIEMYŒL", "BORECZEK", "BUKOWINA", "CIERPISZ", _
     "KLÊCZANY", "KRZYWA", "RUDA", "SZKODNA", _
    "ZAB£OCIE", "ZAGORZYCE", "BRONISZÓW", "BRZEZINY", "GLINIK", "NAWSIE", "RZESZÓW", "PASZCZYNA")
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
                .Execute
            End With
            Do While .Find.Found
                i = i + 1
                .Start = .Words.First.Start
                .End = .Words.First.End
                .MoveEndWhile " ", -1
                .InsertAfter ChrW(9658)
                 .End = .End + 1
                .Font.Color = 204
                .Font.Bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
    'Application.ScreenUpdating = True
    'MsgBox i & " instances found."
    
    'pogrubienie podpisu i dodanie tabulatora
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True

    With Selection.Find
        .Text = "so^p"
        .Replacement.Text = "^tso^p"
        .MatchWildcards = False
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "so ^p"
        .Replacement.Text = "^tso^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
        With Selection.Find
        .Text = "jas^p"
        .Replacement.Text = "^tjas^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "mf^p"
        .Replacement.Text = "^tmf^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
      With Selection.Find
        .Text = "red^p"
        .Replacement.Text = "^tred^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "Miko³aj Froñ^p"
        .Replacement.Text = "^tMiko³aj Froñ^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'pogrubianie akapitu drugiego lub trzeciego
    
    If Documents.Count > 0 Then
        With ActiveDocument
            If .Paragraphs.Count > 2 Then
                If Len(.Paragraphs(2).Range.Text) > 1 Then
                    .Paragraphs(2).Range.Bold = True
                Else
                    .Paragraphs(3).Range.Bold = True
                End If
            Else
                Beep
                MsgBox "Paragraphs 2 and 3 don't exist!"
            End If
        End With
    Else
        Beep
        MsgBox "There is no document open!"
    End If
    

    
    
     'dodanie enter do tesktów s³awka
     
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: ARCHIWUM –"
        .Replacement.Text = "FOTO: ARCHIWUM^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: S. OSKARBSKI –"
        .Replacement.Text = "FOTO: S. OSKARBSKI^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: R. BERDO –"
        .Replacement.Text = "^&^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: ARCHIWUM –"
        .Replacement.Text = "^&^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    

    
    
    'dwuwyrazowe Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .SmallCaps = False
        .AllCaps = False
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "WIELOPOLE SKRZYÑSKIE"
        .Replacement.Text = "WIELOPOLE SKRZYÑSKIE" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "SÊDZISZÓW M£P."
        .Replacement.Text = "SÊDZISZÓW M£P." & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
       
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "BOREK WIELKI"
        .Replacement.Text = "BOREK WIELKI" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "BOREK MA£Y"
        .Replacement.Text = "BOREK MA£Y" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "£¥CZKI KUCHARSKIE"
        .Replacement.Text = "£¥CZKI KUCHARSKIE" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "CZARNA SÊDZISZOWSKA"
        .Replacement.Text = "CZARNA SÊDZISZOWSKA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "WOLICA £UGOWA"
        .Replacement.Text = "WOLICA £UGOWA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "WOLICA PIASKOWA"
        .Replacement.Text = "WOLICA PIASKOWA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "GÓRA ROPCZYCKA"
        .Replacement.Text = "GÓRA ROPCZYCKA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "KAWÊCZYN SÊDZISZOWSKI"
        .Replacement.Text = "KAWÊCZYN SÊDZISZOWSKI" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "WOLA OCIECKA"
        .Replacement.Text = "WOLA OCIECKA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "CA£Y POWIAT"
        .Replacement.Text = "CA£Y POWIAT" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .SmallCaps = False
        .AllCaps = False
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "SÊDZISZÓW MA£OPOLSKI"
        .Replacement.Text = "SÊDZISZÓW MA£OPOLSKI" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    
    
    'formatowanie po formatowaniu
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .SmallCaps = False
        .AllCaps = False
        .Bold = False
        .Color = wdColorAutomatic
    End With
    With Selection.Find
        .Text = "FOTO: UG WIELOPOLE SKRZYÑSKIE" & ChrW(9658)
        .Replacement.Text = "FOTO: UG WIELOPOLE SKRZYÑSKIE"
        .Forward = True
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'Sub convert_paragraph_after_finding()
    Dim oRng As Range
    Set oRng = ActiveDocument.Range
    With oRng.Find
         'Find the word
        Do While .Execute(findText:="FOTO:", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    oRng.End = oRng.Paragraphs(1).Range.End
     'collapse the range to its end
    oRng.Collapse 0
     'move the end of the range to the end of the following paragraph
 'If Len(oRng) = 0 Then
    'Set oRng = Nothing
    'Else
 
    oRng.End = oRng.Next.Paragraphs(1).Range.End
    
    
    
    
     'If the paragraph is empty
    If Len(oRng) = 1 Then
        oRng.Collapse 0
         'move the end of the range to the end of the following paragraph
        oRng.End = oRng.Next.Paragraphs(1).Range.End
    End If
'End If
             'format the range
            With oRng
                .Font.Name = "Times New Roman"
                .Font.Size = 8
                .Font.Italic = wdToggle
                .Bold = False
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 10
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                     
                     'what alignment do you want. You had both?
                     '.Alignment = wdAlignParagraphLeft
                    .Alignment = wdAlignParagraphJustify
                     '
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
                .Collapse 0
            End With
             'and stop looking
            'Exit Do
        Loop
    End With
lbl_Exit:
    Set oRng = Nothing
    'Exit Sub
    
    
     'Sub convert_paragraph_after_finding()
    Dim rRng As Range
    Set rRng = ActiveDocument.Range
    With rRng.Find
         'Find the word
        Do While .Execute(findText:="RAMKA", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    rRng.End = ActiveDocument.Range.End
    If rRng.Paragraphs.Count > 1 Then
    rRng.Start = rRng.Paragraphs(2).Range.Start
    rRng.Select 'for testing only
Else
    MsgBox "There are no more paragraphs!"
End If
             'format the range
            With rRng
                .Font.Name = "Arial Narrow"
                .Font.Size = 8
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                     
                     'what alignment do you want. You had both?
                     '.Alignment = wdAlignParagraphLeft
                    .Alignment = wdAlignParagraphJustify
                     '
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                    
                End With
            End With
             'and stop looking
            Exit Do
        Loop
    End With
lblr_Exit:
    Set orRng = Nothing
    
    'Sub FOTO()
             
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
                .Font.AllCaps = True
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                    .Alignment = wdAlignParagraphJustify
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
                .Collapse 0
            End With
             
            'Exit Do
        Loop
    End With
lblf_Exit:
    Set fRng = Nothing
    Exit Sub
    
    
End Sub


Sub podpissomfjas()
'
' Makro1 Makro
'
With ActiveDocument
    If Len(.Paragraphs(2).Range.Text) > 1 Then
        .Paragraphs(2).Range.Bold = True
    Else
        .Paragraphs(3).Range.Bold = True
    End If
End With


    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True

    With Selection.Find
        .Text = "so^p"
        .Replacement.Text = "^tso^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
        With Selection.Find
        .Text = "jas^p"
        .Replacement.Text = "^tjas^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "mf^p"
        .Replacement.Text = "^tmf^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
      With Selection.Find
        .Text = "red^p"
        .Replacement.Text = "^tred^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   End Sub




Sub bold_ogloszenie()
'
'

   'format_arial
'
    Selection.WholeStory
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With Selection.ParagraphFormat
        .SpaceBefore = 2
        .SpaceBeforeAuto = False
        .SpaceAfter = 2
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
'
'formatowanie numeru og³oszenia

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(^#/^#^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(^#^#/^#^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^#^#/^#^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^#/^#^#^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^#/^#^#^#^#)"
        .Replacement.Text = "^t^&"
        .Forward = True
        .Wrap = wdFindContinue
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^#^#/^#^#^#^#)"
        .Replacement.Text = "^t^&"
        .Forward = True
        .Wrap = wdFindContinue
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do odwo³ania)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^# w ramce)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^# + RAMKA)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^#^# wyt³uszczenie)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^# wyt³uszczenie)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^#^# w ramce)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^#^# + RAMKA)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do ^#^# RAMKA)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do ^#^# RAMKA)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 3
    With Selection.Find
        .Text = "(do nr ^#^# RAMKA)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(do "
        .Replacement.Text = "^t(do "
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
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchByte = False
.MatchAllWordForms = False
.MatchSoundsLike = False
.MatchWildcards = False
.MatchFuzzy = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "^p^p"
.Replacement.Text = "^p"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchByte = False
.MatchAllWordForms = False
.MatchSoundsLike = False
.MatchWildcards = False
.MatchFuzzy = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

'tabulator 3,93 do prawej
'
'
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.93) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        
        
'bold dwóch pierwszych wyrazów
        Dim oPara As Paragraph
    For Each oPara In ActiveDocument.Paragraphs
If Len(oPara.Range.Text) > 1 Then
oPara.Range.Words(1).Font.Bold = True
oPara.Range.Words(2).Font.Bold = True
End If

        
    Next oPara


'Sub RAMkA_oglo()



Dim check As Boolean
Dim search As String
Dim para As Paragraph
Dim tempStr As String
Dim txt As String

search = "RAMKA"

For Each para In ActiveDocument.Paragraphs
    txt = para.Range.Text
    tempStr = (txt)
    check = InStr(tempStr, search)

    If check = True Then
        With para.Range
        With .ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.1)
        .RightIndent = CentimetersToPoints(0.1)
        .SpaceBefore = 6
        .SpaceAfter = 6
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .TextboxTightWrap = wdTightNone
    End With
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.83) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    End If
Next

    
  'Sub bold_wytluszczenie()

Dim oRng As Range
Dim oRngFormat As Range
  Set oRng = ActiveDocument.Range
  With oRng.Find
    Do While .Execute(findText:="wyt³uszczenie", MatchWholeWord:=True)
      oRng.Collapse wdCollapseEnd
      On Error GoTo Err_Handler
      Set oRngFormat = oRng.Paragraphs(1).Range
      With oRngFormat
        .Font.Name = "Arial Narrow"
        .Bold = True
        .Collapse 0
      End With
    Loop
  End With
lbl_Exit:
  Set oRng = Nothing
  Exit Sub
Err_Handler:
  Resume lbl_Exit
  
  
    
  
  
  
End Sub
    

Sub two_bold()
'
' two_bold Makro
'
'

Application.ScreenUpdating = False
Dim i As Long
With ActiveDocument.Range
  With .Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "<[?]{1,}"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = True
    .MatchWildcards = True

  End With
  Do While .Find.Found
    i = i + 1
    .End = .End - 1
    .Style = "Strong"
    .Collapse wdCollapseEnd
    .Find.Execute
  Loop
End With
Application.ScreenUpdating = True
MsgBox i & " definitions bolded."
End Sub

Sub para_bold()


With ActiveDocument
    If Len(.Paragraphs(2).Range.Text) > 1 Then
        .Paragraphs(2).Range.Bold = True
    Else
        .Paragraphs(3).Range.Bold = True
    End If
End With
End Sub

Sub ogloszenie()

Dim oPara As Paragraph
    For Each oPara In ActiveDocument.Paragraphs
        If oPara.Range.Words(1).Font.Bold Then
        oPara.Range.Words(1).Font.Bold = True
        End If
    Next oPara
End Sub
Sub tabulatory()
'
' Makro5 Makro
'
'
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.93) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
End Sub

Sub foto_bezbold()
'
' foto_bezbold Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: ARCHIWUM"
        .Replacement.Text = "FOTO: ARCHIWUM"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: K.IGNAS"
        .Replacement.Text = "FOTO: K.IGNAS"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: UG WIELOPOLE SKRZYÑSKIE"
        .Replacement.Text = "FOTO: UG WIELOPOLE SKRZYÑSKIE"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: UG OSTRÓW"
        .Replacement.Text = "FOTO: UG OSTRÓW"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: UG ROPCZYCE"
        .Replacement.Text = "FOTO: UG ROPCZYCE"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: KPP ROPCZYCE"
        .Replacement.Text = "FOTO: KPP ROPCZYCE"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: CZYTELNIKA"
        .Replacement.Text = "FOTO: CZYTELNIKA"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: ARCHIWUM"
        .Replacement.Text = "FOTO: ARCHIWUM"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: K.IGNAS"
        .Replacement.Text = "FOTO: K.IGNAS"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: M. FROÑ"
        .Replacement.Text = "FOTO: M. FROÑ"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 5
        .Bold = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = "FOTO: D.KIE£EK"
        .Replacement.Text = "FOTO: D.KIE£EK"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
            End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub
Sub dwuwyrazowe()
'
' dwuwyrazowe Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "WIELOPOLE SKRZYÑSKIE"
        .Replacement.Text = "WIELOPOLE SKRZYÑSKIE" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "SÊDZISZÓW M£P."
        .Replacement.Text = "SÊDZISZÓW M£P." & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
       
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "BOREK WIELKI"
        .Replacement.Text = "BOREK WIELKI" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "BOREK MA£Y"
        .Replacement.Text = "BOREK MA£Y" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "£¥CZKI KUCHARSKIE"
        .Replacement.Text = "£¥CZKI KUCHARSKIE" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "CZARNA SÊDZISZOWSKA"
        .Replacement.Text = "CZARNA SÊDZISZOWSKA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "WOLICA £UGOWA"
        .Replacement.Text = "WOLICA £UGOWA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "WOLICA PIASKOWA"
        .Replacement.Text = "WOLICA PIASKOWA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "GÓRA ROPCZYCKA"
        .Replacement.Text = "GÓRA ROPCZYCKA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "KAWÊCZYN SÊDZISZOWSKI"
        .Replacement.Text = "KAWÊCZYN SÊDZISZOWSKI" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "WOLA OCIECKA"
        .Replacement.Text = "WOLA OCIECKA" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "CA£Y POWIAT"
        .Replacement.Text = "CA£Y POWIAT" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 204
    End With
    With Selection.Find
        .Text = "SÊDZISZÓW MA£OPOLSKI"
        .Replacement.Text = "SÊDZISZÓW MA£OPOLSKI" & ChrW(9658)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
           
End Sub


Sub arial_oglo()
'
' arial_oglo Makro
'
'
    Selection.WholeStory
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With Selection.ParagraphFormat
        .SpaceBefore = 2
        .SpaceBeforeAuto = False
        .SpaceAfter = 2
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub


Sub delete_empty_paragraph()
     
     
     
     
   Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "^p^p"
.Replacement.Text = "^p"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchByte = False
.MatchAllWordForms = False
.MatchSoundsLike = False
.MatchWildcards = False
.MatchFuzzy = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub Para2_Bold()
    If Documents.Count > 0 Then
        With ActiveDocument
            If .Paragraphs.Count > 2 Then
                If Len(.Paragraphs(2).Range.Text) > 1 Then
                    .Paragraphs(2).Range.Bold = True
                Else
                    .Paragraphs(3).Range.Bold = True
                End If
            Else
                Beep
                MsgBox "Paragraphs 2 and 3 don't exist!"
            End If
        End With
    Else
        Beep
        MsgBox "There is no document open!"
    End If
End Sub


Sub bold_two_words()

Dim oPara As Paragraph
    For Each oPara In ActiveDocument.Paragraphs
If Len(oPara.Range.Text) > 1 Then
oPara.Range.Words(1).Font.Bold = True
oPara.Range.Words(2).Font.Bold = True
End If

        
    Next oPara
    
    End Sub
    
    'With ActiveDocument
   ' If Len(oPara.Range.Text) > 1 Then
       ' oPara(2).Range.bold = True
   ' Else
     '   .Paragraphs(3).Range.bold = True
  '  End If
'End With


Sub convert_paragraph_after_finding()
    Dim oRng As Range
    Set oRng = ActiveDocument.Range
    With oRng.Find
         'Find the word
        Do While .Execute(findText:="FOTO:", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    oRng.End = oRng.Paragraphs(1).Range.End
     'collapse the range to its end
    oRng.Collapse 0
     'move the end of the range to the end of the following paragraph
    oRng.End = oRng.Next.Paragraphs(1).Range.End
     'If the paragraph is empty
    If Len(oRng) = 1 Then
        oRng.Collapse 0
         'move the end of the range to the end of the following paragraph
        oRng.End = oRng.Next.Paragraphs(1).Range.End
    End If
             'format the range
            With oRng
                .Font.Name = "Times New Roman"
                .Font.Size = 8
                .Font.Italic = wdToggle
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 10
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                     
                     'what alignment do you want. You had both?
                     '.Alignment = wdAlignParagraphLeft
                    .Alignment = wdAlignParagraphJustify
                     '
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
            End With
             'and stop looking
            Exit Do
        Loop
    End With
lbl_Exit:
    Set oRng = Nothing
    Exit Sub
End Sub
Sub prepare_code()

Dim oPara As Paragraph
    For Each oPara In ActiveDocument.Paragraphs
If Len(oPara.Range.Text) > 1 Then
oPara.Range.Words(1).Font.Bold = True
oPara.Range.Words(2).Font.Bold = True
End If

        
    Next oPara
    
    End Sub


Sub tv_puls()
'
'
'
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Rows.Delete
    Selection.Columns.Delete
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.Columns.Delete
    Selection.Columns.Delete
    Selection.Columns.Delete
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^#"
        .Replacement.Text = "(^&)"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ")("
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(0)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "MOCNE SOBOTNIE KINO!"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ZOBACZ TO! BLOK POWTÓRKOWY"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "NIEDZIELA Z GWIAZDAMI"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub

Sub sport_reporter()
'

'formatowanie czcionki i tabulatora

    Selection.WholeStory
    Selection.Font.Name = "Georgia"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    'ActiveDocument.DefaultTabStop = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.93) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces

'pogrubianie nazw miejscowosci dyscyplin
      
      'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("SIATKÓWKA", "SUMO", "TENIS", "BOKS", "SZACHY", "KARATE", "HALOWA PI£KA NO¯NA", _
    "PI£KA NO¯NA", "PODNOSZENIE CIÊ¯ARÓW", "PI£KARSKIE WIEŒCI", "ZAPASY", "KOLARSTWO")
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
                .Execute
            End With
            Do While .Find.Found
                'i = i + 1
                '.Start = .Words.First.Start
                '.End = .Words.First.End
                '.MoveEndWhile " ", -1
                .InsertAfter ChrW(9658)
                 '.End = .End + 1
                .Font.Color = 204
                .Font.Bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
    
    'Application.ScreenUpdating = True
    'MsgBox i & " instances found."
    
    
    
    
    
    
    
    'pogrubienie podpisu i dodanie tabulatora
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True

 
    
    With Selection.Find
        .Text = "so^p"
        .Replacement.Text = "^tso^p"
        .MatchWildcards = False
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
        With Selection.Find
        .Text = "jas^p"
        .Replacement.Text = "^tjas^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "mf^p"
        .Replacement.Text = "^tmf^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
      
      With Selection.Find
        .Text = "red^p"
        .Replacement.Text = "^tred^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      With Selection.Find
        .Text = "Marcin Jastrzêbski^p"
        .Replacement.Text = "^tMarcin Jastrzêbski^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'pogrubianie akapitu drugiego lub trzeciego lub czwartego
    
   
    
    If Documents.Count > 0 Then
        With ActiveDocument
            If .Paragraphs.Count > 2 Then
                If Len(.Paragraphs(2).Range.Text) > 1 Then
                    .Paragraphs(2).Range.Bold = True
                Else
                    .Paragraphs(3).Range.Bold = True
                End If
            Else
                Beep
                MsgBox "Paragraphs 2 and 3 don't exist!"
            End If
        End With
    Else
        Beep
        MsgBox "There is no document open!"
    End If
     
     If Documents.Count > 0 Then
        With ActiveDocument
            If .Paragraphs.Count > 2 Then
                If Len(.Paragraphs(3).Range.Text) > 1 Then
                    .Paragraphs(3).Range.Bold = True
                Else
                    .Paragraphs(4).Range.Bold = True
                End If
             End If
        End With
    End If
                                               
               
    
    'formatowanie po formatowaniu
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .SmallCaps = False
        .AllCaps = False
        .Bold = False
        .Color = 204
    End With
    With Selection.Find
        .Text = "FOTO: UG WIELOPOLE SKRZYÑSKIE" & ChrW(9658)
        .Replacement.Text = "FOTO: UG WIELOPOLE SKRZYÑSKIE"
        .Forward = True
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     'enter po podpisie s³awka
     
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: S. OSKARBSKI –"
        .Replacement.Text = "FOTO: S. OSKARBSKI^p"
        .Forward = True
        '.Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
   
   
   
    'Sub edycja_podpisu_marcina()
    
     Dim qRng As Word.Range
    Set qRng = ActiveDocument.Range
    With qRng.Find
        .Text = "Marcin Jastrzêbski^p791 673 137^p"
        With .Replacement
            '.ClearFormatting
            .Font.Name = "Arial Narrow"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = 1403396
            .Font.Underline = wdUnderlineNone
            With .ParagraphFormat 'the whole paragraph
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(0.9)
                .Alignment = wdAlignParagraphRight
                .LineUnitBefore = 0
                .LineUnitAfter = 0
            End With
        End With
        .Execute Replace:=wdReplaceAll
    End With
lblq_Exit:
    Set qRng = Nothing
   

Dim zRng As Word.Range
    Set zRng = ActiveDocument.Range
    With zRng.Find
        .Text = "m.jastrzebski@reportergazeta.pl"
        
        With .Replacement
            .ClearFormatting
            .Font.Name = "Arial Narrow"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = 1403396
            .Font.Underline = wdUnderlineNone
            
            With .ParagraphFormat 'the whole paragraph
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(0.9)
                .Alignment = wdAlignParagraphRight
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                
            End With
        End With
        .Execute Replace:=wdReplaceAll
    End With
lblu_Exit:
    Set zRng = Nothing
    

    ' powrot_na_poczatek_dokumentu
    
        
    Selection.WholeStory
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.MoveDown Unit:=wdLine, Count:=3
    


   
     'Sub convert_paragraph_after_finding()
    Dim oRng As Range
    Set oRng = ActiveDocument.Range
    With oRng.Find
         'Find the word
        Do While .Execute(findText:="RAMKA", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    oRng.End = ActiveDocument.Range.End
    If oRng.Paragraphs.Count > 1 Then
    oRng.Start = oRng.Paragraphs(2).Range.Start
    oRng.Select 'for testing only
Else
    MsgBox "There are no more paragraphs!"
End If
             'format the range
            With oRng
                .Font.Name = "Arial Narrow"
                .Font.Size = 8
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                     
                     'what alignment do you want. You had both?
                     '.Alignment = wdAlignParagraphLeft
                    .Alignment = wdAlignParagraphJustify
                     '
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                    .TabStops.ClearAll
                    .TabStops.Add Position:=CentimetersToPoints(3), _
                    Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
                    .TabStops.Add Position:=CentimetersToPoints(3.3), _
                    Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
                    .TabStops.Add Position:=CentimetersToPoints(3.88), _
                    Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
                End With
            End With
             'and stop looking
            Exit Do
        Loop
    End With
lbl_Exit:
    Set oRng = Nothing
    
    
      
   ' powrot_na_poczatek_dokumentu
    
        
    Selection.WholeStory
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.MoveDown Unit:=wdLine, Count:=3
      
    
    
    'Sub convert_paragraph_after_finding()
    Dim xRng As Range
    Set xRng = ActiveDocument.Range
    With xRng.Find
         'Find the word
        Do While .Execute(findText:="FOTO:", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    xRng.End = xRng.Paragraphs(1).Range.End
     'collapse the range to its end
    xRng.Collapse 0
     'move the end of the range to the end of the following paragraph
    xRng.End = xRng.Next.Paragraphs(1).Range.End
     'If the paragraph is empty
    If Len(xRng) = 1 Then
        xRng.Collapse 0
         'move the end of the range to the end of the following paragraph
        xRng.End = xRng.Next.Paragraphs(1).Range.End
    End If
             'format the range
            With xRng
                .Font.Name = "Times New Roman"
                .Font.Size = 8
                .Font.Italic = wdToggle
                .Bold = False
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                     
                     'what alignment do you want. You had both?
                     '.Alignment = wdAlignParagraphLeft
                    .Alignment = wdAlignParagraphJustify
                     '
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
                .Collapse 0
            End With
             'and stop looking
            'Exit Do
        Loop
    End With
lbl1_Exit:
    Set xRng = Nothing
    
    
           'formatowanie podpisu
      
      Dim jRng As Word.Range
    Set jRng = ActiveDocument.Range
    With jRng.Find
        .Text = "FOTO: M. JASTRZÊBSKI"
        
        With .Replacement
            .ClearFormatting
            .Font.Name = "Times New Roman"
            .Font.Size = 5
            
            
            
            
            With .ParagraphFormat 'the whole paragraph
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(0.9)
                
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                
            End With
        End With
        .Execute Replace:=wdReplaceAll
    End With
lblj_Exit:
    Set jRng = Nothing
     
     'Sub FOTO()
             
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
                .Font.AllCaps = True
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                    .Alignment = wdAlignParagraphJustify
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
                .Collapse 0
            End With
             
            'Exit Do
        Loop
    End With
lblf_Exit:
    Set fRng = Nothing
    'Exit Sub

'pogrubianie nazw miejscowosci dyscyplin
      
      'Application.ScreenUpdating = False
    Dim q As Long, w As Long, ArrFnds()
    ArrFnds = Array("ROZMOWA REPORTERA", "FLESZEM", "OPINIA", "ZDANIEM TRENERA", "ZDANIEM ZAWODNIKA", "P£YWANIE")
    For q = 0 To UBound(ArrFnds)
        With ActiveDocument.Range
            With .Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ArrFnds(q)
                .Highlight = False
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = True
                .Execute
            End With
            Do While .Find.Found
                'w = i + 1
                '.Start = .Words.First.Start
                '.End = .Words.First.End
                '.MoveEndWhile " ", -1
                .InsertAfter ChrW(9660)
                 '.End = .End + 1
                .Font.Color = 204
                .Font.Bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
    
    'Application.ScreenUpdating = True
    'MsgBox i & " instances found."
    
    'Sub srodtytul()

   
Dim cPar As Paragraph
    For Each cPar In ActiveDocument.Range.Paragraphs
       
            If cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify _
               And cPar.Range.Font.Size = 9 _
                                And cPar.Range.Font.Name = "Georgia" _
                 And Len(cPar.Range) > 10 _
                 And Len(cPar.Range) < 60 _
                 And Not cPar.Range.Characters.last = "." Then
             
                cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                cPar.Range.InsertBefore Chr(13)
                cPar.Range.Font.Size = 10
                cPar.Range.Font.Bold = True
              
            
                'And cPar.Range.Font.bold = True
            End If
       
    Next

End Sub


 Sub convert_paragraph_after_finding_sport()
    Dim sRng As Range
    Set sRng = ActiveDocument.Range
    With sRng.Find
         'Find the word
        Do While .Execute(findText:="RAMKA", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    sRng.End = sRng.Paragraphs(1).Range.End
     'collapse the range to its end
    sRng.Collapse 0
     'move the end of the range to the end of the following paragraph
    sRng.End = sRng.Next.Paragraphs(1).Range.End
     'If the paragraph is empty
    If Len(sRng) = 1 Then
        sRng.Collapse 0
         'move the end of the range to the end of the following paragraph
        sRng.End = sRng.Next.Paragraphs(1).Range.End
    End If
             'format the range
            With sRng
                .Font.Name = "Arial Narrow"
                .Font.Size = 8
                .Font.Italic = wdToggle
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 10
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                     
                     'what alignment do you want. You had both?
                     '.Alignment = wdAlignParagraphLeft
                    .Alignment = wdAlignParagraphJustify
                     '
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
            End With
             'and stop looking
            Exit Do
        Loop
    End With
lbl_Exit:
    Set sRng = Nothing
    Exit Sub
End Sub
Sub sport_test()
'
 'Sub convert_paragraph_after_finding()
    Dim oRng As Range
    Set oRng = ActiveDocument.Range
    With oRng.Find
         'Find the word
        Do While .Execute(findText:="RAMKA", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    oRng.End = ActiveDocument.Range.End
    If oRng.Paragraphs.Count > 1 Then
    oRng.Start = oRng.Paragraphs(2).Range.Start
    oRng.Select 'for testing only
Else
    MsgBox "There are no more paragraphs!"
End If
             'format the range
            With oRng
                .Font.Name = "Arial Narrow"
                .Font.Size = 8
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                     
                     'what alignment do you want. You had both?
                     '.Alignment = wdAlignParagraphLeft
                    .Alignment = wdAlignParagraphJustify
                     '
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
            End With
             'and stop looking
            Exit Do
        Loop
    End With
lbl_Exit:
    Set oRng = Nothing
    
    
    'Sub convert_paragraph_after_finding()
    Dim xRng As Range
    Set xRng = ActiveDocument.Range
    With xRng.Find
         'Find the word
        Do While .Execute(findText:="FOTO:", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    xRng.End = xRng.Paragraphs(1).Range.End
     'collapse the range to its end
    xRng.Collapse 0
     'move the end of the range to the end of the following paragraph
    xRng.End = xRng.Next.Paragraphs(1).Range.End
     'If the paragraph is empty
    If Len(xRng) = 1 Then
        xRng.Collapse 0
         'move the end of the range to the end of the following paragraph
        xRng.End = xRng.Next.Paragraphs(1).Range.End
    End If
             'format the range
            With xRng
                .Font.Name = "Times New Roman"
                .Font.Size = 8
                .Font.Italic = wdToggle
                .Bold = False
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                     
                     'what alignment do you want. You had both?
                     '.Alignment = wdAlignParagraphLeft
                    .Alignment = wdAlignParagraphJustify
                     '
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
            End With
             'and stop looking
            Exit Do
        Loop
    End With
lbl1_Exit:
    Set xRng = Nothing
    Exit Sub
    
    'Sub edycja_podpisu_marcina()
    
    Dim qRng As Word.Range
    Set qRng = ActiveDocument.Range
    With qRng.Find
        .Text = "^tMarcin Jastrzêbski^p791 673 137^p"
        With .Replacement
            '.ClearFormatting
            .Font.Name = "Arial Narrow"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = 1403396
            With .ParagraphFormat 'the whole paragraph
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(0.9)
                .Alignment = wdAlignParagraphRight
                .LineUnitBefore = 0
                .LineUnitAfter = 0
            End With
        End With
        .Execute Replace:=wdReplaceAll
    End With
lblq_Exit:
    Set qRng = Nothing
   

Dim zRng As Word.Range
    Set zRng = ActiveDocument.Range
    With zRng.Find
        .Text = "m.jastrzebski@reportergazeta.pl"
        
        With .Replacement
            .ClearFormatting
            .Font.Name = "Arial Narrow"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = 1403396
            .Font.Underline = wdUnderlineNone
            
            With .ParagraphFormat 'the whole paragraph
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(0.9)
                .Alignment = wdAlignParagraphRight
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                
            End With
        End With
        .Execute Replace:=wdReplaceAll
    End With
lblu_Exit:
    Set zRng = Nothing
    Exit Sub
    
    End Sub
    
    Sub edycja_podpisu_marcina()
 
 Dim qRng As Word.Range
    Set qRng = ActiveDocument.Range
    With qRng.Find
        .Text = "Marcin Jastrzêbski^p791 673 137^p"
        With .Replacement
            '.ClearFormatting
            .Font.Name = "Arial Narrow"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = 1403396
            With .ParagraphFormat 'the whole paragraph
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(0.9)
                .Alignment = wdAlignParagraphRight
                .LineUnitBefore = 0
                .LineUnitAfter = 0
            End With
        End With
        .Execute Replace:=wdReplaceAll
    End With
lblq_Exit:
    Set qRng = Nothing
   

Dim zRng As Word.Range
    Set zRng = ActiveDocument.Range
    With zRng.Find
        .Text = "m.jastrzebski@reportergazeta.pl"
        
        With .Replacement
            .ClearFormatting
            .Font.Name = "Arial Narrow"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = 1403396
            .Font.Underline = wdUnderlineNone
            
            With .ParagraphFormat 'the whole paragraph
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(0.9)
                .Alignment = wdAlignParagraphRight
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                
            End With
        End With
        .Execute Replace:=wdReplaceAll
    End With
lbl_Exit:
    Set zRng = Nothing
    Exit Sub
    
    
   
End Sub
    
    
 
Sub krzyzowka()
'
' Makro4 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "^#^#)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
       
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Poziomo:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Pionowo:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = "Bogdan Witek"
        .Replacement.Text = "^t^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
       
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With Selection.Find
        .Text = "."
        .Replacement.Text = ";"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
       
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub
Sub Makro1()
'
'pogrubianie nazw miejscowosci dyscyplin
      
      'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("PI£KA NO¯NA")
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
                .Execute
            End With
            Do While .Find.Found
                i = i + 1
                .Start = .Words.First.Start
                .End = .Words.First.End
                .MoveEndWhile " ", -1
                .InsertAfter ChrW(9658)
                 .End = .End + 1
                .Font.Color = 204
                .Font.Bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
End Sub
Sub Makro2()
'
'  Sub insert_bold_after_second_word()
      'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("PI£KA NO¯NA", "SÊDZISZÓW MA£OPOLSKI", "WIELOPOLE SKRZYÑSKIE", "SÊDZISZÓW M£P.", _
    "BOREK WIELKI", "BOREK MA£Y", "£¥CZKI KUCHARSKIE", _
    "CZARNA SÊDZISZOWSKA", "WOLICA £UGOWA", "WOLICA PIASKOWA", "GÓRA ROPCZYCKA", _
    "KAWÊCZYN SÊDZISZOWSKI", "WOLA OCIECKA", "CA£Y POWIAT")
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
                .Execute
            End With
            Do While .Find.Found
                'i = i + 1
                '.Start = .Words.First.Start
                '.End = .Words.First.End
                '.MoveEndWhile " ", -1
                .InsertAfter ChrW(9658)
                 '.End = .End + 1
                .Font.Color = 204
                .Font.Bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Exit Do
            Loop
            
        End With
    Next
End Sub



Sub test_strzalek()
'
'  Sub insert_bold_after_second_word()
      'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("PI£KA NO¯NA", "SÊDZISZÓW MA£OPOLSKI", "WIELOPOLE SKRZYÑSKIE", "SÊDZISZÓW M£P.", _
    "BOREK WIELKI", "BOREK MA£Y", "£¥CZKI KUCHARSKIE", _
    "CZARNA SÊDZISZOWSKA", "WOLICA £UGOWA", "WOLICA PIASKOWA", "GÓRA ROPCZYCKA", _
    "KAWÊCZYN SÊDZISZOWSKI", "WOLA OCIECKA", "CA£Y POWIAT", _
    "ROPCZYCE", "IWIERZYCE", "BÊDZIENICA", "BYSTRZYCA", "NOCKOWA", _
    "OLCHOWA", "OLIMPÓW", "SIELEC", "WIERCANY", "WIŒNIOWA", "BRZEZÓWKA", "GNOJNICA", "LUBZINA", _
    "MA£A", "NIEDWIADA", "OKONIN", "BLIZNA", "KAMIONKA", "KOZODRZA", "OCIEKA", "OSTRÓW", "SKRZYSZÓW", _
    "ZD¯ARY", "BÊDZIEMYŒL", "BORECZEK", "BUKOWINA", "CIERPISZ", _
     "KLÊCZANY", "KRZYWA", "RUDA", "SZKODNA", _
    "ZAB£OCIE", "ZAGORZYCZE", "BRONISZÓW", "BRZEZINY", "GLINIK", "NAWSIE", "RZESZÓW", "PASZCZYNA")
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
                .Execute
            End With
            Do While .Find.Found
                'i = i + 1
                '.Start = .Words.First.Start
                '.End = .Words.First.End
                '.MoveEndWhile " ", -1
                .InsertAfter ChrW(9658)
                 '.End = .End + 1
                .Font.Color = 204
                .Font.Bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Exit Do
            Loop
            
        End With
    Next
End Sub

Sub pogrubienia_do_znakow_tvp_i_inne_2wersja_kodu()
Dim oRng As Range
    Set oRng = ActiveDocument.Range
    oRng.Font.Bold = True
    With oRng.Find
        Do While .Execute(findText:=" (")
            oRng.End = oRng.Paragraphs(1).Range.End - 1
            oRng.Start = oRng.Start + 1
            oRng.Font.Bold = False
            oRng.Collapse 0
        Loop
    End With
    Set oRng = ActiveDocument.Range
    With oRng.Find
        Do While .Execute(findText:=" - ")
            oRng.End = oRng.Paragraphs(1).Range.End - 1
            oRng.Start = oRng.Start + 1
            oRng.Font.Bold = False
            oRng.Collapse 0
        Loop
    End With
lbl_Exit:
    Set oRng = Nothing
    Exit Sub
End Sub
Sub pogrubienia_do_znakow_tvp_i_inne()
    Dim r As Range
    
    Set r = ActiveDocument.Range
    
    r.Font.Bold = True

    With r.Find
        .MatchWildcards = True
        
        .Text = "[\(\-]*^13"
        .Replacement.Font.Bold = False
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub
Sub Makro4()
'
' Makro4 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Raport smogowy – wiem czym oddycham"
        .Replacement.Text = "Raport smogowy"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub program_formatowanie()
'
' program_formatowanie_czcionka_wciecie Makro
'
'
    Selection.WholeStory
    Selection.Font.Name = "Arial Narrow"
    Selection.Font.Size = 8
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .FirstLineIndent = CentimetersToPoints(-0.66)
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
    End With
    
    Dim r As Range
    
    Set r = ActiveDocument.Range
    
    r.Font.Bold = True

    With r.Find
        .MatchWildcards = True
        
        .Text = "[\(\-]*^13"
        .Replacement.Font.Bold = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub array_test_test()
'

'formatowanie czcionki i tabulatora

    Selection.WholeStory
    Selection.Font.Name = "Georgia"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.95)
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = False
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.93) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces

'pogrubianie nazw miejscowosci
      
      'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("ROPCZYCE", "IWIERZYCE", "OSTRÓW", "BÊDZIENICA", "BYSTRZYCA", "NOCKOWA", _
    "OLCHOWA", "OLIMPÓW", "SIELEC", "WIERCANY", "WIŒNIOWA", "BRZEZÓWKA", "GNOJNICA", "LUBZINA", _
    "MA£A", "NIEDWIADA", "OKONIN", "BLIZNA", "KAMIONKA", "KOZODRZA", "OCIEKA", "SKRZYSZÓW", _
    "ZD¯ARY", "BÊDZIEMYŒL", "BORECZEK", "BUKOWINA", "CIERPISZ", "KAWÊCZYN", _
     "KLÊCZANY", "KRZYWA", "RUDA", "SZKODNA", _
    "ZAB£OCIE", "ZAGORZYCE", "BRONISZÓW", "BRZEZINY", "GLINIK", "NAWSIE", "RZESZÓW", "PASZCZYNA", _
    "WIELOPOLE SKRZYÑSKIE", "SÊDZISZÓW M£P.", "BOREK WIELKI", "BOREK MA£Y", _
"£¥CZKI KUCHARSKIE", "CZARNA SÊDZISZOWSKA", "WOLICA £UGOWA", "WOLICA PIASKOWA", _
"GÓRA ROPCZYCKA", "KAWÊCZYN SÊDZISZOWSKI", "WOLA OCIECKA", "CA£Y POWIAT", "SÊDZISZÓW MA£OPOLSKI", "WIELOPOLE SKRZ.")
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
                .Execute
            End With
            Do While .Find.Found
                'i = i + 1
                '.Start = .Words.First.Start
                '.End = .Words.First.End
                '.MoveEndWhile " ", -1
                .InsertAfter ChrW(9658)
                 '.End = .End + 1
                .Font.Color = 204
                .Font.Bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
    'Application.ScreenUpdating = True
    'MsgBox i & " instances found."
    
    'pogrubienie podpisu i dodanie tabulatora
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True

    With Selection.Find
        .Text = "so^p"
        .Replacement.Text = "^tso^p"
        .MatchWildcards = False
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "so ^p"
        .Replacement.Text = "^tso^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
        With Selection.Find
        .Text = "jas^p"
        .Replacement.Text = "^tjas^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "dc^p"
        .Replacement.Text = "^tdc^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "naj^p"
        .Replacement.Text = "^tnaj^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "kbr^p"
        .Replacement.Text = "^tkbr^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "mf^p"
        .Replacement.Text = "^tmf^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
      With Selection.Find
        .Text = "red^p"
        .Replacement.Text = "^tred^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "Miko³aj Froñ^p"
        .Replacement.Text = "^tMiko³aj Froñ^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "pga^p"
        .Replacement.Text = "^tpga^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "mp^p"
        .Replacement.Text = "^tmp^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "pg^p"
        .Replacement.Text = "^tpg^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "ol^p"
        .Replacement.Text = "^tol^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'pogrubianie akapitu drugiego lub trzeciego
    
    If Documents.Count > 0 Then
        With ActiveDocument
            If .Paragraphs.Count > 2 Then
                If Len(.Paragraphs(2).Range.Text) > 1 Then
                    .Paragraphs(2).Range.Bold = True
                Else
                    .Paragraphs(3).Range.Bold = True
                End If
            Else
                Beep
                MsgBox "Paragraphs 2 and 3 don't exist!"
            End If
        End With
    Else
        Beep
        MsgBox "There is no document open!"
    End If
    

    
    
     'dodanie enter do tesktów s³awka
     
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: ARCHIWUM –"
        .Replacement.Text = "FOTO: ARCHIWUM^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: S. OSKARBSKI –"
        .Replacement.Text = "FOTO: S. OSKARBSKI^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: R. BERDO –"
        .Replacement.Text = "^&^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "FOTO: ARCHIWUM –"
        .Replacement.Text = "^&^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    

        
       
    
    
    'formatowanie po formatowaniu
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .SmallCaps = False
        .AllCaps = False
        .Bold = False
        .Color = wdColorAutomatic
    End With
    With Selection.Find
        .Text = "FOTO: UG WIELOPOLE SKRZYÑSKIE" & ChrW(9658)
        .Replacement.Text = "FOTO: UG WIELOPOLE SKRZYÑSKIE"
        .Forward = True
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .SmallCaps = False
        .AllCaps = False
        .Bold = False
        .Color = wdColorAutomatic
    End With
    With Selection.Find
        .Text = "FOTO: KPP ROPCZYCE" & ChrW(9658)
        .Replacement.Text = "FOTO: KPP ROPCZYCE"
        .Forward = True
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .SmallCaps = False
        .AllCaps = False
        .Bold = False
        .Color = wdColorAutomatic
    End With
    With Selection.Find
        .Text = "FOTO: UG OSTRÓW" & ChrW(9658)
        .Replacement.Text = "FOTO: UG OSTRÓW"
        .Forward = True
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
        
    
     'Sub convert_paragraph_after_finding()
    Dim rRng As Range
    Set rRng = ActiveDocument.Range
    With rRng.Find
         'Find the word
        Do While .Execute(findText:="RAMKA", MatchWholeWord:=True)
     'move the end of the range to the end of the paragraph containing the found word
    rRng.End = ActiveDocument.Range.End
    If rRng.Paragraphs.Count > 1 Then
    rRng.Start = rRng.Paragraphs(2).Range.Start
    rRng.Select 'for testing only
Else
    MsgBox "There are no more paragraphs!"
End If
             'format the range
            With rRng
                .Font.Name = "Arial Narrow"
                .Font.Size = 8
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                    .Alignment = wdAlignParagraphJustify
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                    
                End With
            End With
             'and stop looking
            Exit Do
        Loop
    End With
lblr_Exit:
    Set orRng = Nothing
    
    'Sub FOTO()
             
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
                .Font.AllCaps = True
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                    .Alignment = wdAlignParagraphJustify
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
                .Collapse 0
            End With
             
            'Exit Do
        Loop
    End With
lblf_Exit:
    Set fRng = Nothing
    'Exit Sub
    
    
    'Sub srodtytul()

   
Dim cPar As Paragraph
    For Each cPar In ActiveDocument.Range.Paragraphs
       
            If cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify _
               And cPar.Range.Font.Size = 9 _
                                And cPar.Range.Font.Name = "Georgia" _
                 And Len(cPar.Range) > 10 _
                 And Len(cPar.Range) < 60 _
                 And Not cPar.Range.Characters.last = "." Then
             
                cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                cPar.Range.InsertBefore Chr(13)
                cPar.Range.Font.Size = 10
                cPar.Range.Font.Bold = True
              
            
                'And cPar.Range.Font.bold = True
            End If
       
    Next
    

    
    'Sub ScratchMacro()
'A basic Word macro coded by Greg Maxey, http://gregmaxey.com/word_tips.html, 4/5/2018
Dim oRng As Range
Dim oRngFormat As Range
  Set oRng = ActiveDocument.Range
  With oRng.Find
    Do While .Execute(findText:="FOTO", MatchWholeWord:=True)
      oRng.Collapse wdCollapseEnd
      On Error GoTo Err_Handler
      Set oRngFormat = oRng.Paragraphs(1).Next.Range
      'If oRngFormat = 0 Then
    ' Set oRngFormat = orng.Paragraphs(2).Next.Range
    ' End If
     
      With oRngFormat
        .Font.Name = "Times New Roman"
        .Font.Size = 8
        .Font.Italic = wdToggle
        .Bold = False
        With .ParagraphFormat
          .LeftIndent = CentimetersToPoints(0)
          .RightIndent = CentimetersToPoints(0)
          .SpaceBefore = 0
          .SpaceBeforeAuto = False
          .SpaceAfter = 10
          .SpaceAfterAuto = False
          .LineSpacingRule = wdLineSpaceMultiple
          .LineSpacing = LinesToPoints(0.9)
          .Alignment = wdAlignParagraphLeft
          .WidowControl = True
          .KeepWithNext = False
          .KeepTogether = False
          .PageBreakBefore = False
          .NoLineNumber = False
          .Hyphenation = True
          .FirstLineIndent = CentimetersToPoints(0)
          .OutlineLevel = wdOutlineLevelBodyText
          .CharacterUnitLeftIndent = 0
          .CharacterUnitRightIndent = 0
          .CharacterUnitFirstLineIndent = 0
          .LineUnitBefore = 0
          .LineUnitAfter = 0
          .MirrorIndents = False
          .TextboxTightWrap = wdTightNone
        End With
        .Collapse 0
      End With
    Loop
  End With
lbl_Exit:
  Set oRng = Nothing
  Exit Sub
Err_Handler:
  Resume lbl_Exit
  
     
  
End Sub
    
    


Sub ScratchMacro()
'A basic Word macro coded by Greg Maxey, http://gregmaxey.com/word_tips.html, 4/5/2018
Dim oRng As Range
Dim oRngFormat As Range
  Set oRng = ActiveDocument.Range
  With oRng.Find
    Do While .Execute(findText:="FOTO", MatchWholeWord:=True)
      oRng.Collapse wdCollapseEnd
      On Error GoTo Err_Handler
      Set oRngFormat = oRng.Paragraphs(1).Next.Range
      With oRngFormat
        .Font.Name = "Times New Roman"
        .Font.Size = 8
        .Font.Italic = wdToggle
        .Bold = False
        With .ParagraphFormat
          .LeftIndent = CentimetersToPoints(0)
          .RightIndent = CentimetersToPoints(0)
          .SpaceBefore = 0
          .SpaceBeforeAuto = False
          .SpaceAfter = 10
          .SpaceAfterAuto = False
          .LineSpacingRule = wdLineSpaceMultiple
          .LineSpacing = LinesToPoints(0.9)
          .Alignment = wdAlignParagraphLeft
          .WidowControl = True
          .KeepWithNext = False
          .KeepTogether = False
          .PageBreakBefore = False
          .NoLineNumber = False
          .Hyphenation = True
          .FirstLineIndent = CentimetersToPoints(0)
          .OutlineLevel = wdOutlineLevelBodyText
          .CharacterUnitLeftIndent = 0
          .CharacterUnitRightIndent = 0
          .CharacterUnitFirstLineIndent = 0
          .LineUnitBefore = 0
          .LineUnitAfter = 0
          .MirrorIndents = False
          .TextboxTightWrap = wdTightNone
        End With
        .Collapse 0
      End With
    Loop
  End With
lbl_Exit:
  Set oRng = Nothing
  Exit Sub
Err_Handler:
  Resume lbl_Exit
End Sub

 Sub test_ramki()


 Dim oRng As Range
    Set oRng = ActiveDocument.Range
    With oRng.Find
    Do While .Execute(findText:="RAMKA", MatchWholeWord:=True)
    oRng.End = ActiveDocument.Range.End
    If oRng.Paragraphs.Count > 1 Then
    oRng.Start = oRng.Paragraphs(2).Range.Start
    oRng.Select

End If
              With oRng
                .Font.Name = "Arial Narrow"
                .Font.Size = 8
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                    .Alignment = wdAlignParagraphJustify
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                    .TabStops.ClearAll
                    .TabStops.Add Position:=CentimetersToPoints(3), _
                    Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
                    .TabStops.Add Position:=CentimetersToPoints(3.3), _
                    Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
                    .TabStops.Add Position:=CentimetersToPoints(3.88), _
                    Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
                End With
            End With
            Exit Do
        Loop
    End With
lbl_Exit:
    Set oRng = Nothing
End Sub


Sub test_bold_wytluszczenie()


 'Sub ScratchMacro()

Dim oRng As Range
Dim oRngFormat As Range
  Set oRng = ActiveDocument.Range
  With oRng.Find
    Do While .Execute(findText:="wyt³uszczenie", MatchWholeWord:=True)
      oRng.Collapse wdCollapseEnd
      On Error GoTo Err_Handler
      Set oRngFormat = oRng.Paragraphs(1).Range
      With oRngFormat
        .Font.Name = "Arial Narrow"
        .Bold = True
        .Collapse 0
      End With
    Loop
  End With
lbl_Exit:
  Set oRng = Nothing
  Exit Sub
Err_Handler:
  Resume lbl_Exit
End Sub


Sub test_oglo_w_ramce()

Dim xRng As Range
Dim xRngFormat As Range
  Set xRng = ActiveDocument.Range
  With xRng.Find
    Do While .Execute(findText:="ramce", MatchWholeWord:=True)
      xRng.Collapse wdCollapseEnd
      On Error GoTo Err_Handler
      Set xRngFormat = xRng.Paragraphs(1).Range
      With xRngFormat
        '.Font.Name = "Arial Narrow"
        '.Font.Size = 8
        '.Font.Italic = wdToggle
        '.bold = True
       With .ParagraphFormat
          .LeftIndent = CentimetersToPoints(0.1)
          .RightIndent = CentimetersToPoints(0.1)
          .SpaceBefore = 6
          .SpaceBeforeAuto = False
          .SpaceAfter = 6
          .SpaceAfterAuto = False
          .LineSpacingRule = wdLineSpaceMultiple
          .LineSpacing = LinesToPoints(0.9)
          .Alignment = wdAlignParagraphJustify
          .WidowControl = True
          .KeepWithNext = False
          .KeepTogether = False
          .PageBreakBefore = False
          .NoLineNumber = False
          .Hyphenation = True
          .FirstLineIndent = CentimetersToPoints(0)
          .OutlineLevel = wdOutlineLevelBodyText
          .CharacterUnitLeftIndent = 0
          .CharacterUnitRightIndent = 0
          .CharacterUnitFirstLineIndent = 0
          .LineUnitBefore = 0
          .LineUnitAfter = 0
          .MirrorIndents = False
          .TabStops.ClearAll
          .TabStops.Add Position:=CentimetersToPoints(3.88), _
        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
          .TextboxTightWrap = wdTightNone
        End With
        .Collapse 0
      End With
    Loop
  End With
lbl_Exit:
  Set oRng = Nothing
  Exit Sub
Err_Handler:
  Resume lbl_Exit
End Sub
Sub Makro3()
'
' Makro3 Makro
'
'
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.HomeKey Unit:=wdLine
    Selection.MoveRight Unit:=wdCharacter, Count:=38, Extend:=wdExtend
    Selection.Range.Case = wdTitleWord
End Sub
Sub subtitle()
'
' subtitle Makro
'
'<Subtitle></Subtitle>

    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 10
        .Bold = True
    End With
    With Selection.Find.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
    End With
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = "<Subtitle>^&</Subtitle>^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p</Subtitle>"
        .Replacement.Text = "</Subtitle>"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '<Body></Body>
   
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 9
        .Bold = False
        .Italic = False
    End With
    With Selection.Find.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
    End With
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = "<Body>^&</Body>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
         End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '<Description></Description>
   
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 8
        .Bold = False
        .Italic = True
    End With
    'With Selection.Find.ParagraphFormat
        '.Alignment = wdAlignParagraphJustify
    'End With
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = "<Description>^&</Description>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
         End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
End Sub

Sub Makro5()
'
' Makro5 Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 10
        .Bold = False
        .Italic = False
    End With
    With Selection.Find.ParagraphFormat
        '.SpaceBeforeAuto = False
        '.SpaceAfterAuto = False
        .Alignment = wdAlignParagraphJustify
    End With
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = "<Body>^&</Body>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
         End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub




 Sub FOTO_xml()
             
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
                .Font.AllCaps = True
                .InsertAfter "</Credits>"
                .InsertBefore "<Credits>"
                With .ParagraphFormat
                    
                End With
                .Collapse 0
            End With
             
            'Exit Do
        Loop
    End With
lblf_Exit:
    Set fRng = Nothing
    Exit Sub
    
    
End Sub


Sub Title_xml()
Dim oRng As Range
  Set oRng = ActiveDocument.Range
  With oRng.Find
    .Text = "*^13"
    .Format = True
    .MatchWildcards = True
    Do While .Execute
      If oRng.Font.Size > 14 Then
        oRng.InsertAfter "</Title>"
        oRng.InsertBefore "<Title>"
      End If
    Loop
  End With
End Sub

Sub remove_empty_paragraph_xml()

    
Dim oRng As Range
  Set oRng = ActiveDocument.Range
  With oRng.Find
    .Text = "*^13"
    .Format = True
    .MatchWildcards = True
    Do While .Execute
      If oRng.Font.Size > 14 And Len(oRng) > 1 Then
        oRng.End = oRng.End - 1
        oRng.InsertAfter "</Title>"
        oRng.InsertBefore "<Title>"
        oRng.Collapse wdCollapseEnd
        oRng.End = oRng.End + 1
      End If
    Loop
  End With


End Sub

Sub vignette_xml()
             
    Dim vRng As Range
    Set vRng = ActiveDocument.Range
    With vRng.Find
         
        Do While .Execute(findText:=ChrW(9660), MatchWholeWord:=True)
     
    vRng.Start = vRng.Paragraphs(1).Range.Start
     
            With vRng
                .InsertAfter "</Vignette>"
                .InsertBefore "<Vignette>"
                
                With .ParagraphFormat
                    
                End With
                .Collapse 0
            End With
                  
        Loop
    End With
lblf_Exit:
    Set vRng = Nothing
    Exit Sub
    
    
End Sub



Sub vignette_sport_xml()
             
    Dim vRng As Range
    Set vRng = ActiveDocument.Range
    With vRng.Find
         
        Do While .Execute(findText:=ChrW(9658), MatchWholeWord:=True)
                With .ParagraphFormat
                    .Alignment = wdAlignParagraphCenter
                End With
    
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
    Exit Sub
    
    
End Sub


Sub lead_xml()
             
               
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
    Exit Sub
    
    
End Sub

Sub Publico_xml()

 ' usuwanie spacji bia³ych
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^s"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    'Sub podpisy_xml()

    
     With Selection.Find
        .Text = "^tso"
        .Replacement.Text = "^p<Author>S³awomir Oskarbski</Author>^p</Story></Article>"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tmf"
        .Replacement.Text = "^p<Author>Miko³aj Froñ</Author>^p</Story></Article>"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tjas"
        .Replacement.Text = "^p<Author>Marcin Jastrzêbski</Author>^p</Story></Article>"
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
   
   
   
    
    
     '<Body></Body>felieton
   
    'Sub italic_body_loop_xml()
          
    Dim iRng As Range
    Set iRng = ActiveDocument.Range
    With iRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                End With
                 With .Font
                    .Size = 9
                    .Italic = True
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    iRng.Start = iRng.Paragraphs(1).Range.Start
              With iRng
                .InsertBefore "<Body>"
            End With
            
         iRng.End = iRng.Paragraphs(1).Range.End - 1
             With iRng
                .InsertAfter "</Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lbli_Exit:
    Set iRng = Nothing
   'Exit Sub
    
    
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
                .InsertBefore "^p<Description>"
            End With
            
         dRng.End = dRng.Paragraphs(1).Range.End - 1
             With dRng
                .InsertAfter "</Description>^p</Picture>"
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
    .Replacement.Text = "<Article>^p<Story>^p<Title>\1</Title>^p"
    .MatchWildcards = True
    .Wrap = wdFindContinue
    For i = 29 To 144
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
                .Font.AllCaps = True
                .InsertAfter "</Credits>"
                .InsertBefore "<Picture>^p<Credits>"
                With .ParagraphFormat
                    
                End With
                .Collapse 0
            End With
             
            'Exit Do
        Loop
    End With
lblx_Exit:
    Set fRng = Nothing
    'Exit Sub
    
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
    
    
    'Sub body_bold_xml()
Dim bxPar As Paragraph
Dim bxRng As Range
    For Each bxPar In ActiveDocument.Range.Paragraphs
        Set bxRng = bxPar.Range
        bxRng.End = bxRng.End - 1
        If bxRng.Font.Bold = True _
           And bxRng.Font.Size = 9 _
           And Not bxRng.Characters.First = "<" _
           And Len(bxRng) > 1 Then
            bxRng.InsertAfter "</Body>"
            bxRng.InsertBefore "<Body>"
        End If
    Next bxPar
    Set bxPar = Nothing
    Set bxRng = Nothing

    'Sub horoskop_xml()
Dim hPar As Paragraph
Dim hRng As Range
    For Each hPar In ActiveDocument.Range.Paragraphs
        Set hRng = hPar.Range
        hRng.End = hRng.End - 1
        If hRng.Font.Bold = True _
           And hRng.Font.Size = 8 _
           And Len(hRng) > 1 _
           And hRng.Font.Color = RGB(0, 109, 53) Then
            hRng.InsertAfter "</Subtitle>"
            hRng.InsertBefore "<Subtitle originalStyle=""horoskop_znak"">"
        End If
    Next hPar
    Set hPar = Nothing
    Set hRng = Nothing
    
    
  'Sub add_xml_intro_outro1()


ActiveDocument.Content.InsertBefore "<?xml version='1.0' encoding='UTF-8' standalone='no'?>" & Chr(13) & "<Root>" & Chr(13)
ActiveDocument.Content.InsertAfter Chr(13) & "</Root>"


End Sub


Sub titlte_xml()
Application.ScreenUpdating = False
Dim i As Long
With ActiveDocument.Range
  With .Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Forward = True
    .Format = True
    .Text = "(*)^13"
    .Replacement.Text = "<Title>\1</Title>^p"
    .MatchWildcards = True
    .Wrap = wdFindContinue
    For i = 29 To 144
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
End Sub

Sub Makro6()
'
' usuwanie spacji bia³ych
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^s"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



 Sub subtitle_loop_xml()
          
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
                .InsertBefore "<Lead>"
            End With
            
         sRng.End = sRng.Paragraphs(1).Range.End - 1
             With sRng
                .InsertAfter "</Lead>"
                .Collapse 0
            End With
         
        Loop
    End With
lbls_Exit:
    Set sRng = Nothing
   Exit Sub
    
    End Sub
    
    Sub body_loop_xml()
          
    Dim bRng As Range
    Set bRng = ActiveDocument.Range
    With bRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphJustify
                End With
                 With .Font
                    .Size = 9
                    .Bold = False
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
   Exit Sub
   
   
   
    
    End Sub
    
    Sub podpisy1_xml()

    
     With Selection.Find
        .Text = "^tso"
        .Replacement.Text = "^p<Author>S³awomir Oskarbski</Author>^p"
        
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tmf"
        .Replacement.Text = "^p<Author>Miko³aj Froñ</Author>^p"
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tjas"
        .Replacement.Text = "^p<Author>Marcin Jastrzêbski</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    End Sub
    
    
     Sub italic_body_loop_xml()
          
    Dim iRng As Range
    Set iRng = ActiveDocument.Range
    With iRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                End With
                 With .Font
                    .Size = 9
                    .Italic = True
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    iRng.Start = iRng.Paragraphs(1).Range.Start
              With iRng
                .InsertBefore "<Body>"
            End With
            
         iRng.End = iRng.Paragraphs(1).Range.End - 1
             With iRng
                .InsertAfter "</Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lbli_Exit:
    Set iRng = Nothing
   Exit Sub
    
    End Sub
    
    
    Sub description_loop_xml()
          
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
   Exit Sub
    
    End Sub
    
    Sub xml_rozmowa()

    
    'Sub subtitle_loop_xml()
          
    Dim sRng As Range
    Set sRng = ActiveDocument.Range
    With sRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                    
                End With
                 With .Font
                    .Size = 9
                    .Bold = True
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    sRng.Start = sRng.Paragraphs(1).Range.Start
              With sRng
                .InsertBefore "<Body><![CDATA[{body:bold}"
            End With
            
         sRng.End = sRng.Paragraphs(1).Range.End - 1
             With sRng
                .InsertAfter "{/body:bold}<br />"
                .Collapse 0
            End With
         
        Loop
    End With
lbls_Exit:
    Set sRng = Nothing
   'Exit Sub
   
   'podpis
   Dim pRng As Range
    Set pRng = ActiveDocument.Range
    With pRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphRight
                    
                End With
                 With .Font
                    .Size = 9
                    .Bold = True
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    pRng.Start = pRng.Paragraphs(1).Range.Start
              With pRng
                .InsertBefore "<Body>{body:bold}"
            End With
            
         pRng.End = pRng.Paragraphs(1).Range.End - 1
             With pRng
                .InsertAfter "{/body:bold}</Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lblp_Exit:
    Set pRng = Nothing
   'Exit Sub
   
   
   
   'Sub lead_rozmowa_xml()
          
    Dim lRng As Range
    Set lRng = ActiveDocument.Range
    With lRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphJustify
                End With
                 With .Font
                    .Size = 9
                    .Bold = True
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
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
    For i = 29 To 144
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

'Sub body_loop_xml()
          
    Dim bRng As Range
    Set bRng = ActiveDocument.Range
    With bRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                End With
                 With .Font
                    .Size = 9
                    .Bold = False
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                
         bRng.End = bRng.Paragraphs(1).Range.End - 1
             With bRng
                .InsertAfter "<br /><br />]]></Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lblb_Exit:
    Set bRng = Nothing
   'Exit Sub


'wstawianie do znaczników kiedy nie ma pe³nego formatowania paragrafu
   
               
Dim sPara As Paragraph
Dim sRnga As Range
    For Each sPara In ActiveDocument.Range.Paragraphs
        Set sRnga = sPara.Range
        sRnga.End = sRnga.End - 1
        If Len(sRnga) > 1 _
            And sRnga.Font.Size = 9 _
            And sRnga.ParagraphFormat.Alignment = wdAlignParagraphLeft _
            And sRnga.Characters.First.Bold = True _
            And Not sRnga.Characters.First = "<" _
           And Not sRnga.Characters.last = ">" Then
            
            sRnga.InsertAfter "<br /><br />]]></Body>"
         
        End If
    Next sPara
    Set sPara = Nothing
    Set sRnga = Nothing




'Sub vignette_xml()
             
    Dim vRng As Range
    Set vRng = ActiveDocument.Range
    With vRng.Find
         
        Do While .Execute(findText:=ChrW(9660), MatchWholeWord:=True)
     
    vRng.Start = vRng.Paragraphs(1).Range.Start
     
            With vRng
                .InsertAfter "</Vignette>"
                .InsertBefore "<Vignette>"
                
                With .ParagraphFormat
                    
                End With
                .Collapse 0
            End With
                  
        Loop
    End With
lblf_Exit:
    Set vRng = Nothing
    'Exit Sub
    
    
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
                .InsertAfter "</Credits>" & Chr(13) & "</Picture>"
                .InsertBefore "<Picture>" & Chr(13) & _
                "<Image href=""file://images/nazwa_zdjecia.jpg""></Image>" & Chr(13) & _
                "<Credits>" & Chr(13)
                                  
              
                .Collapse 0
            End With
            
    
            'Exit Do
        Loop
    End With
    
lblx_Exit:
    Set fRng = Nothing
    'Exit Sub
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</Picture><Description>"
        .Replacement.Text = "<Description>"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll




 
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
                .InsertAfter "</Description>^p</Picture>"
                .Collapse 0
            End With
         
        Loop
    End With
lbld_Exit:
    Set dRng = Nothing
   'Exit Sub

   
   'Sub add_xml_intro_outro1()


ActiveDocument.Content.InsertBefore "<?xml version='1.0' encoding='UTF-8' standalone='no'?>" & Chr(13) & "<Root>" & Chr(13)
ActiveDocument.Content.InsertAfter Chr(13) & "</Story></Article></Root>"



   
    
    End Sub
    
Sub add_xml_intro_outro()

ActiveDocument.Content.InsertBefore "<?xml version='1.0' encoding='UTF-8' standalone='no'?>" & Chr(13) & "<Root>" & Chr(13)
ActiveDocument.Content.InsertAfter Chr(13) & "</Root>"

End Sub

 
    
    
    
    Sub ogloszenia_xml()
    
    
     'Sub oglo_naglowki()



    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "RÓ¯NE: ^p^tSPRZEDAM^p^tODDAM"
        .Replacement.Text = "RÓ¯NE: SPRZEDAM, ODDAM"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "NIERUCHOMOŒCI^p^tSPRZEDAM^p^tMIESZKANIE:"
        .Replacement.Text = "NIERUCHOMOŒCI: SPRZEDAM MIESZKANIE:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "NIERUCHOMOŒCI^p^tSPRZEDAM^p^tDOM:"
        .Replacement.Text = "NIERUCHOMOŒCI: SPRZEDAM DOM:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "NIERUCHOMOŒCI^p^tKUPIÊ"
        .Replacement.Text = "NIERUCHOMOŒCI KUPIÊ:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "NIERUCHOMOŒCI^p^tWYNAJMÊ"
        .Replacement.Text = "NIERUCHOMOŒCI WYNAJMÊ:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "US£UGI^pOGÓLNOBUDOWLANE:"
        .Replacement.Text = "US£UGI OGÓLNOBUDOWLANE:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "US£UGI^p^tRÓ¯NE"
        .Replacement.Text = "US£UGI RÓ¯NE"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "PRACA^p^tDAM"
        .Replacement.Text = "PRACA DAM"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "PRACA^p^tSZUKAM"
        .Replacement.Text = "PRACA SZUKAM"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'End Sub

     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t ^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     
     
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "^p"
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
    
    'Sub add_after_bold()


  'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("ROPCZYmmmCE")
    For x = 0 To UBound(ArrFnd)
        With ActiveDocument.Range
            With .Find
                .ClearFormatting
                .Replacement.ClearFormatting
                '.Text = ArrFnd(x)
                '.Highlight = False
                '.Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = True
                .Font.Bold = True
                .Font.Size = 9
                .Font.Italic = False
                '.Font.Color = RGB(33, 33, 32)
                .Execute
            End With
            Do While .Find.Found
                'i = i + 1
                '.Start = .Words.First.Start
               ' .End = .Words.First.End
                '.MoveEndWhile " ", -1
                .InsertAfter "{/body:bold}"
                 .End = .End
                '.Font.Color = 204
                '.Font.bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
    'Application.ScreenUpdating = True
    'MsgBox i & " instances found."'Sub body_loop_xml()
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        '.Font.Color = RGB(33, 33, 32)
        '.Font.Size = 3
        .Text = "^p{/body:bold}"
        .Replacement.Text = " {/body:bold}^p"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
          
    Dim bRng As Range
    Set bRng = ActiveDocument.Range
    With bRng.Find
        
                 'With .ParagraphFormat
                    '.Alignment = wdAlignParagraphLeft
               'End With
                 With .Font
                    .Size = 9
                    '.bold = False
                    .Italic = False
                    '.Color = RGB(33, 33, 32)
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                
         bRng.End = bRng.Paragraphs(1).Range.End - 1
             With bRng
                .InsertBefore "<Body>{body:bold}"
                .InsertAfter "</Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lblb_Exit:
    Set bRng = Nothing
   'Exit Sub
          
   
   'Makro7 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        '.Font.Color = RGB(33, 33, 32)
        .Font.Size = 3
        .Text = "("
        .Replacement.Text = "<Body originalStyle=""nr_oglo"">("
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        '.Font.Color = RGB(33, 33, 32)
        .Font.Size = 3
        .Text = ")"
        .Replacement.Text = ")</Body>"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'Sub ramka_oglo_xml()

Dim check As Boolean
Dim search As String
Dim para As Paragraph
Dim tempStr As String
Dim txt As String

search = "RAMKA"

For Each para In ActiveDocument.Paragraphs
    txt = para.Range.Text
    tempStr = (txt)
    check = InStr(tempStr, search)

    If check = True Then
        If Not para.Range = ActiveDocument.Range.Paragraphs.First.Range Then
              If Len(para.Range) > 2 _
               And para.Range.Font.Name = "Arial Narrow" Then
                para.Previous(1).Range.InsertBefore "<Body originalStyle=""ramka_oglo"">"
                para.Previous(1).Range.Characters.last.InsertBefore "</Body>"
                End If
                End If
    End If
Next


'Sub arial_center_modul()

   
Dim mPar As Paragraph
    For Each mPar In ActiveDocument.Range.Paragraphs
       
            If mPar.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter _
               And mPar.Range.Font.Size = 10 _
               And Len(mPar.Range) > 2 _
               And mPar.Range.Font.Name = "Arial Narrow" Then
                mPar.Range.InsertBefore "<Body originalStyle=""oglo_modul"">"
                mPar.Range.Characters.last.InsertBefore "</Body>"
              
            
               
            End If
       
    Next
       
   
    
    
    'Sub oglo_naglowki_xml()
'And hRng.Font.Color = RGB(255, 255, 255) = True hRng.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
Dim hPar As Paragraph
Dim hRng As Range
    For Each hPar In ActiveDocument.Range.Paragraphs
        Set hRng = hPar.Range
        hRng.End = hRng.End - 1
        If Len(hRng) > 1 And hRng.Font.Color = RGB(255, 255, 254) Or hRng.Font.Color = RGB(255, 250, 246) Or hRng.Font.Color = RGB(255, 255, 255) Then
            hRng.InsertAfter "</Subtitle><Body><![CDATA[<br />]]></Body>"
            hRng.InsertBefore "<Subtitle originalStyle=""oglo_tyt"">"
        End If
    Next hPar
    Set hPar = Nothing
    Set hRng = Nothing
    
    
    'Sub zdjecie_oglo_xml()

Dim checkz As Boolean
Dim searchz As String
Dim paraz As Paragraph
Dim tempStrz As String
Dim txtz As String

search = "ZDJÊCIE"

For Each paraz In ActiveDocument.Paragraphs
    txtz = paraz.Range.Text
    tempStr = (txtz)
    checkz = InStr(tempStr, search)

    If checkz = True Then
        If Not paraz.Range = ActiveDocument.Range.Paragraphs.First.Range Then
              If Len(paraz.Range) > 2 _
               And paraz.Range.Font.Name = "Arial Narrow" Then
                paraz.Previous(1).Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)

                'para.Previous(1).Range.Characters.last.InsertBefore "</Body>"
                End If
                End If
    End If
Next
    
    
    
    'Sub add_xml_intro_outro()

ActiveDocument.Content.InsertBefore "<?xml version='1.0' encoding='UTF-8' standalone='no'?>" & Chr(13) & _
"<Root>" & Chr(13) & "<Article>" & Chr(13) & "<Story>" & Chr(13) & "<Title>Og³oszenia drobne</Title>"
ActiveDocument.Content.InsertAfter Chr(13) & "</Story>" & Chr(13) & "</Article>" & Chr(13) & "</Root>"

'{/<Body>{body:bold}body:bold}</Body>
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "{/<Body>{body:bold}body:bold}</Body>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'Sub highlight_xml()


Dim iPar As Paragraph
    For Each iPar In ActiveDocument.Range.Paragraphs
       
            If Not iPar.Range.Characters.First = "<" _
            And Not iPar.Range.Characters.last = ">" _
            Then
                iPar.Range.HighlightColorIndex = wdBrightGreen
              
            
                       End If
       
    Next

    
    End Sub
    
    
   
Sub Makro7()
'
' Makro7 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "</Body><Body originalStyle=""nr_oglo"">"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub xml_tabela()
'
' Makro7 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "</td><td>"
        .Forward = True
        .Wrap = wdFindContinue
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'Sub body_loop_xml()
          
    Dim bRng As Range
    Set bRng = ActiveDocument.Range
    With bRng.Find
        
                
                 With .Font
                    .Size = 8
                    .Bold = False
                    .Italic = False
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                
         bRng.End = bRng.Paragraphs(1).Range.End - 1
             With bRng
                .InsertBefore "<tr><td>"
                .InsertAfter "</td></tr>"
                .Collapse 0
            End With
         
        Loop
    End With
lblb_Exit:
    Set bRng = Nothing
  
          
   
   
End Sub

Sub add_after_bold()


  'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("ROPCZYCE")
    For x = 0 To UBound(ArrFnd)
        With ActiveDocument.Range
            With .Find
                .ClearFormatting
                .Replacement.ClearFormatting
                '.Text = ArrFnd(x)
                '.Highlight = False
                '.Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = True
                .Font.Bold = True
                .Execute
            End With
            Do While .Find.Found
                'i = i + 1
                '.Start = .Words.First.Start
               ' .End = .Words.First.End
                '.MoveEndWhile " ", -1
                .InsertAfter "</Bold>"
                 .End = .End
                '.Font.Color = 204
                '.Font.bold = True
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
    'Application.ScreenUpdating = True
    'MsgBox i & " instances found."
    
    End Sub
    
    
      Sub tabela_find()

  
Dim bRng As Range
Set bRng = ActiveDocument.Range
With bRng.Find
        Do While .Execute(findText = "1.*^132.*^13", MatchWholeWord:=True, MatchWildcards:=True)
    
bRng.End = ActiveDocument.Range.End
'bRng.Start = bRng.Paragraphs(1).Range.Start
    With bRng.Find
            .Text = "^t"
        .Replacement.Text = "</td><td>"
        .Forward = True
        .Wrap = wdFindContinue
    End With
        Loop
        Exit Do
  
  End With
               
lblb_Exit:
    Set orRng = Nothing
    
         
    
        
        End Sub
        
        
        Sub xml_tabela1()
'
' Makro7 Makro
'
'
 
    
    
    
    Dim aRng As Range
    Set aRng = ActiveDocument.Range
    With aRng.Find
        
                
                 With .Font
                    .Size = 8
                    '.bold = False
                    '.Italic = False
                    .Underline = wdUnderlineSingle
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                         
             With aRng.Find
                .Text = "^t"
                .Replacement.Text = "</td><td>"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
         
        Loop
    End With
lbla_Exit:
    Set aRng = Nothing
        
    
    
  
          
   
   
End Sub


 Sub add_before_xml()
          
    Dim bRng As Range
    Set bRng = ActiveDocument.Range
    With bRng.Find
        
                'Selection.Find.Replacement.ClearFormatting
                 With .Font
                    
                    .Size = 8
                    '.bold = False
                    '.Italic = False
                    .Underline = wdUnderlineSingle
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                
         bRng.End = bRng.Paragraphs(1).Range.End - 1
             With bRng
                .InsertBefore "<tr><td>"
                .InsertAfter "</td></tr>"
                .Collapse 0
            End With
         
        Loop
    End With
lblb_Exit:
    Set bRng = Nothing
  
          
   
   
End Sub


Sub Makro9()
'
' Makro9 Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = False
        .Italic = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
End Sub


Sub adddd_before_xml()
          
    Dim bRng As Range
    Set bRng = ActiveDocument.Range
    With bRng.Find
        
                
                 With .Font
                    Selection.Find.Replacement.ClearFormatting
                    .Size = 8
                    .Underline = wdUnderlineSingle
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                
         bRng.End = bRng.Paragraphs(1).Range.End - 1
             With bRng
                .InsertBefore "<tr><td>"
                .InsertAfter "</td></tr>"
                .Collapse 0
            End With
         
        Loop
    End With
lblb_Exit:
    Set bRng = Nothing

End Sub

Sub add1_before_xml()
Dim bRng As Range
Dim oPar As Paragraph
  Set bRng = ActiveDocument.Range
  With bRng.Find
    With .Font
      .Size = 8
      .Underline = wdUnderlineSingle
    End With
    Do While .Execute(MatchWholeWord:=True)
      bRng.End = bRng.Paragraphs(1).Range.End - 1
      With bRng
        .Select
        .InsertBefore "<tr><td>"
        .InsertAfter "</td></tr>"
        .Collapse 0
      End With
    Loop
  End With
  For Each oPar In ActiveDocument.Range.Paragraphs
    If oPar.Range.Characters.First = "-" And Not InStr(oPar.Range.Text, "<tr>") > 0 _
    And oPar.Range.Font.Size = 8 Then
      oPar.Range.InsertBefore "<tr><td>"
      oPar.Range.InsertAfter "</td></tr>"
    End If
  Next
lblb_Exit:
  Set bRng = Nothing
End Sub


Sub dodaj_znacznik_tabeli_xml()
Dim bRng As Range
Dim oPar As Paragraph
  Set bRng = ActiveDocument.Range
  With bRng.Find
      With .Font
      .Size = 8
      .Underline = wdUnderlineSingle
       End With
       With bRng.Find
       .Text = "<tr><td>1."
       End With
 
    
    Do While .Execute(MatchWholeWord:=True)
      bRng.End = bRng.Paragraphs(1).Range.End - 1
      With bRng
        .Select
        .InsertBefore "znacznik_tabela"
        .Collapse 0
      End With
    Loop
  End With
  
lblb_Exit:
  Set bRng = Nothing
End Sub


Sub add_tabela_znacznik_after_xml()


Dim oPar As Paragraph
  Set bPar = ActiveDocument.Range
  
  With bPar.Find
    With .Font
    .Size = 8
    .Underline = wdUnderlineSingle
    End With
    
    
 
  For Each oPar In ActiveDocument.Range.Paragraphs
    If oPar.Range.Characters.First = "<" _
    And Not InStrRev(oPar.Range.Text, "-") > 0 _
    And Not InStr(oPar.Next.Range.Text, "<tr>") > 0 Then
      'oPar.Range.InsertBefore "<tr><td>"
      oPar.Range.InsertAfter "znaczek zakonczenia tabeli"
      '.Collapse 0
      On Error Resume Next
    End If
  Next
  
  End With
lblb_Exit:
  'Set bRng = Nothing
End Sub

Sub add_sign_xml()


Dim oPar As Paragraph

 
         


    For Each oPar In ActiveDocument.Range.Paragraphs
    If oPar.Range.Characters.First = "<" _
    And Not InStrRev(oPar.Range.Text, "-") > 0 _
    And Not InStr(oPar.Next.Range.Text, "<tr>") > 0 _
    And Not Len(oPar.Next.Range) = 1 Then
            oPar.Range.InsertAfter "add"
            

    
    End If
  
Next


  
  
  

End Sub


Sub add_sign_xml111()



Dim bPar As Range
Dim last As Long
Dim myparas As Paragraph
 
Set oPar = ActiveDocument.Paragraphs
Set bPar = ActiveDocument.Range
 
last = ActiveDocument.Paragraphs.Count





     For x = 1 To last
     If ActiveDocument.Paragraphs(x).Range.Characters(1) = "<" And Not InStrRev(ActiveDocument.Paragraphs(x).Range.Text, "-") > 0 _
    And Not InStr(ActiveDocument.Paragraphs(x + 1).Range.Text, "<tr>") > 0 Then
     ActiveDocument.Paragraphs(x).Range.InsertAfter "</tbody></table></Body>" & vbCr
  

    End If
    

   Next x

   
End Sub
 
    
Sub add_sign_xml22()
Dim oPar As Paragraph
    For Each oPar In ActiveDocument.Range.Paragraphs
        If Not oPar.Range = ActiveDocument.Range.Paragraphs.last.Range Then
            If oPar.Range.Characters.First = "<" _
               And Not InStrRev(oPar.Range.Text, "-") = 1 _
               And Not InStr(oPar.Next(1).Range.Text, "<tr>") > 0 Then
                oPar.Range.InsertAfter "add"
            End If
        End If
    Next
End Sub


Sub tabela_komplet()


'Sub add1_before_xml()
Dim bRng As Range
Dim oPar As Paragraph
  Set bRng = ActiveDocument.Range
  With bRng.Find
    With .Font
      .Size = 8
      .Underline = wdUnderlineSingle
    End With
    Do While .Execute(MatchWholeWord:=True)
      bRng.End = bRng.Paragraphs(1).Range.End - 1
      With bRng
        .Select
        .InsertBefore "<tr><td>"
        .InsertAfter "</td></tr>"
        .Collapse 0
      End With
    Loop
  End With
  For Each oPar In ActiveDocument.Range.Paragraphs
    If oPar.Range.Characters.First = "-" And Not InStr(oPar.Range.Text, "<tr>") > 0 _
    And oPar.Range.Font.Size = 8 Then
      oPar.Range.InsertBefore "<tr><td>"
      oPar.Range.InsertAfter "</td></tr>"
    End If
  Next
lblb_Exit:
  Set bRng = Nothing
  
  'tabela_tabulator
  
   Dim aRng As Range
    Set aRng = ActiveDocument.Range
    With aRng.Find
        
                
                 With .Font
                    .Size = 8
                    '.bold = False
                    '.Italic = False
                    .Underline = wdUnderlineSingle
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                         
             With aRng.Find
                .Text = "^t"
                .Replacement.Text = "</td><td>"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
         
        Loop
    End With
lbla_Exit:
    Set aRng = Nothing


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
        .InsertBefore "<Body><table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:250px;""><tbody>"
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
               And Not InStr(sPar.Next(1).Range.Text, "<tr>") > 0 Then
                
                
                sPar.Range.InsertAfter "</tbody></table></Body>"
            End If
        End If
    Next

End Sub


Sub add_sign_xml111111()

End Sub



Dim bPar As Range
Dim last As Long
Dim myparas As Paragraph
 
Set oPar = ActiveDocument.Paragraphs
Set bPar = ActiveDocument.Range
 
last = ActiveDocument.Paragraphs.Count





     For x = 1 To last
     If ActiveDocument.Paragraphs(x).Range.Characters(1) = "<" And Not InStrRev(ActiveDocument.Paragraphs(x).Range.Text, "-") > 0 _
    And Not InStr(ActiveDocument.Paragraphs(x + 1).Range.Text, "<tr>") > 0 Then
     ActiveDocument.Paragraphs(x).Range.InsertAfter "</tbody></table></Body>" & vbCr
  

    End If
    

   Next x
   End Sub
   
   Sub add_sign_xml223()
Dim sPar As Paragraph
    For Each sPar In ActiveDocument.Range.Paragraphs
        If Not sPar.Range = ActiveDocument.Range.Paragraphs.last.Range Then
            If sPar.Range.Characters.First = "<" _
                And sPar.Range.Underline = wdUnderlineSingle _
               And Not InStrRev(sPar.Range.Text, "-") > 0 _
               And Not InStr(sPar.Next(1).Range.Text, "<tr>") > 0 Then
                sPar.Range.InsertAfter "</tbody></table></Body>"
            End If
        End If
    Next

End Sub



Sub Publico_xml_sport()

 ' usuwanie spacji bia³ych
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
    .Text = "Miko³aj Froñ"
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
        .Text = "Marcin Jastrzêbski"
        .Replacement.Text = "^p<Author>Marcin Jastrzêbski</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tso"
        .Replacement.Text = "^p<Author>S³awomir Oskarbski</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "^tdc"
        .Replacement.Text = "^p<Author>Dominika Czy¿</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tmf"
        .Replacement.Text = "^p<Author>Miko³aj Froñ</Author>^p"
        .Wrap = wdFindContinue
         End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
       
        .Text = "^tjas"
        .Replacement.Text = "^p<Author>Marcin Jastrzêbski</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
     With Selection.Find
    .Text = "^ttab"
        .Replacement.Text = "^p<Author>W³adys³aw Tabasz</Author>"
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
        .Replacement.Text = "^p<Author>Pawe³ Bochenek</Author>"
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
        .Replacement.Text = "^p<Author>Dominika Czy¿</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      
     With Selection.Find
    .Text = "Ortalion"
        .Replacement.Text = "^p<Author>Bronis³aw Róg</Author>"
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
                .InsertAfter "</Description>^p</Picture>"
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
                .InsertBefore "<Picture>" & Chr(13) & _
                "<Image href=""file://images/nazwa_zdjecia.jpg""></Image>" & Chr(13) & _
                "<Credits>" & Chr(13)
                                  
              
                .Collapse 0
            End With
            
    
            'Exit Do
        Loop
    End With
    
lblx_Exit:
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
        .InsertBefore "<Body><![CDATA[<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:250px;""><tbody>"
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
           And hRng.Font.Size = 8 _
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
    ArrFnd = Array("ROPCZYCE", "IWIERZYCE", "OSTRÓW", "BÊDZIENICA", "BYSTRZYCA", "NOCKOWA", _
    "OLCHOWA", "OLIMPÓW", "SIELEC", "WIERCANY", "WIŒNIOWA", "BRZEZÓWKA", "GNOJNICA", "LUBZINA", _
    "MA£A", "NIEDWIADA", "OKONIN", "BLIZNA", "KAMIONKA", "KOZODRZA", "OCIEKA", "SKRZYSZÓW", _
    "ZD¯ARY", "BÊDZIEMYŒL", "BORECZEK", "BUKOWINA", "CIERPISZ", "KAWÊCZYN", _
     "KLÊCZANY", "KRZYWA", "RUDA", "SZKODNA", "TARNÓW", _
    "ZAB£OCIE", "ZAGORZYCE", "BRONISZÓW", "BRZEZINY", "GLINIK", "NAWSIE", "RZESZÓW", "PASZCZYNA", _
    "WIELOPOLE SKRZYÑSKIE", "SÊDZISZÓW M£P.", "BOREK WIELKI", "BOREK MA£Y", "WARSZAWA", _
"£¥CZKI KUCHARSKIE", "CZARNA SÊDZISZOWSKA", "WOLICA £UGOWA", "WOLICA PIASKOWA", _
"GÓRA ROPCZYCKA", "KAWÊCZYN SÊDZISZOWSKI", "WOLA OCIECKA", "CA£Y POWIAT", "SÊDZISZÓW MA£OPOLSKI", _
"SIATKÓWKA", "SUMO", "TENIS STO£OWY", "BOKS", "SZACHY", "KARATE", "HALOWA PI£KA NO¯NA", _
    "PI£KA NO¯NA", "PODNOSZENIE CIÊ¯ARÓW", "PI£KARSKIE WIEŒCI", "ZAPASY", "KOLARSTWO")
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
        .Text = "Prosimy traktowaæ poni¿szy horoskop z przymru¿eniem oka, gdy¿ prawdopodobieñstwo, ¿e opisane sytuacje kiedykolwiek zaistniej¹ jest znikome i zale¿y jedynie od przypadku."
        .Replacement.Text = "<Body originalStyle=""dowcipy"">^&</Body><Body><![CDATA[<br />]]></Body>"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Wyra¿ane przez Czytelników opinie nie s¹ stanowiskiem redakcji Reporter Gazety."
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
        .Text = "<Vignette>REPORTER I PIENI¥DZE" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "^&^p<Author>Pawe³ Bochenek</Author> "
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
        .Text = "<Vignette>(TROCHÊ) M£ODSZYM OKIEM" & ChrW(9658) & " Miko³aj Froñ </Vignette>"
        .Replacement.Text = "^&^p<Author>Miko³aj Froñ</Author> "
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>CA£KIEM (NIE) OBIEKTYWNIE" & ChrW(9658) & " Wojciech Naja </Vignette>"
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




Sub body_bold()
   
            
    Dim bbRng As Range
    Set bbRng = ActiveDocument.Range
    With bbRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphJustify
                    
                End With
                 With .Font
                    .Size = 9
                    .Bold = True
                    .Name = "Georgia"
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    bbRng.Start = bbRng.Paragraphs(1).Range.Start
              With bbRng
                .InsertBefore "<Body>"
            End With
            
         bbRng.End = bbRng.Paragraphs(1).Range.End - 1
             With bbRng
                .InsertAfter "</Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lblbb_Exit:
    Set bbRng = Nothing
   Exit Sub
   End Sub
   
   
   
   Sub bodybodybold()

   
Dim lPar As Paragraph
    For Each lPar In ActiveDocument.Range.Paragraphs
       
            If lPar.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify _
               And lPar.Range.Font.Bold = True _
               And lPar.Range.Font.Size = 9 _
               And Len(lPar.Range) > 2 Then
                lPar.Range.InsertBefore "<Body>"
                lPar.Range.Characters.last.InsertBefore "</Body>"
              
            
               
            End If
       
    Next
    End Sub
   
   
   Sub body_arial()
   
            
    Dim aRng As Range
    Set aRng = ActiveDocument.Range
    With aRng.Find
        
                 'With .ParagraphFormat
                   ' .Alignment = wdAlignParagraphJustify
               ' End With
                 With .Font
                    .Size = 8
                    '.bold = True
                    .Name = "Arial Narrow"
                    .Underline = wdUnderlineNone
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
    aRng.Start = aRng.Paragraphs(1).Range.Start
              With aRng
                .InsertBefore "<Body originalStyle=""body_ramka"">"
            End With
            
         aRng.End = aRng.Paragraphs(1).Range.End - 1
             With aRng
                .InsertAfter "</Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lblbb_Exit:
    Set aRng = Nothing
   Exit Sub
   End Sub
   
   
   Sub arial_center()

   
Dim cPar As Paragraph
    For Each cPar In ActiveDocument.Range.Paragraphs
       
            If cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter _
               And cPar.Range.Font.Size = 10 _
               And Len(cPar.Range) > 2 _
               And cPar.Range.Font.Name = "Arial Narrow" Then
                cPar.Range.InsertBefore "<Body originalStyle=""oglo_modul"">"
                cPar.Range.Characters.last.InsertBefore "</Body>"
              
            
               
            End If
       
    Next
        End Sub

Sub Makro8()
'
' usuwanie <Root>^p</Story>^p</Article>
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Root>^p</Story>^p</Article>"
        .Replacement.Text = "<Root>"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub ramka_oglo_xml()
oz
Dim check As Boolean
Dim search As String
Dim para As Paragraph
Dim tempStr As String
Dim txt As String

search = "RAMKA"

For Each para In ActiveDocument.Paragraphs
    txt = para.Range.Text
    tempStr = (txt)
    check = InStr(tempStr, search)

    If check = True Then
        If Not para.Range = ActiveDocument.Range.Paragraphs.First.Range Then
              If Len(para.Range) > 2 _
               And para.Range.Font.Name = "Arial Narrow" Then
                para.Previous(1).Range.InsertBefore "<Body originalStyle=""ramka_oglo"">"
                para.Previous(1).Range.Characters.last.InsertBefore "</Body>"
                End If
                End If
    End If
Next
End Sub

Sub test_xml()


Dim lPar As Paragraph
    For Each lPar In ActiveDocument.Range.Paragraphs
       
            If lPar.Range.Font.Bold = True _
               And lPar.Range.Font.Size = 9 _
               And Len(lPar.Range) > 2 Then
               
                'lPar.Range.Characters.last.InsertBefore "</Body>"
              lPar.Range.End = lPar.Range.Characters.last.End - 3
                lPar.Range.InsertAfter "</Body>"
            End If
    Next
    End Sub

Sub Makro10()
'
' Makro10 Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 8
        .Color = 3501312
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([!^13]@)(^13)"
        .Replacement.Text = _
            "<Subtitle originalStyle=""horoskop_znak"">\1</Subtitle>\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub horoskop_xml()
Dim hPar As Paragraph
Dim hRng As Range
    For Each hPar In ActiveDocument.Range.Paragraphs
        Set hRng = hPar.Range
        hRng.End = hRng.End - 1
        If hRng.Font.Bold = True _
           And hRng.Font.Size = 8 _
           And Len(hRng) > 1 _
           And hRng.Font.Color = RGB(0, 109, 53) Then
            hRng.InsertAfter "</Subtitle>"
            hRng.InsertBefore "<Subtitle originalStyle=""horoskop_znak"">"
        End If
    Next hPar
    Set hPar = Nothing
    Set hRng = Nothing
End Sub


Sub dowcipy_xml()
Dim dPar As Paragraph
Dim dRng As Range
    For Each dPar In ActiveDocument.Range.Paragraphs
        Set dRng = dPar.Range
        dRng.End = dRng.End - 1
        If dRng.Font.Size = 9 _
           And dRng.Font.Italic = True _
           And dRng.ParagraphFormat.Alignment = wdAlignParagraphCenter _
           And Len(dRng) > 1 _
            Then
            dRng.InsertAfter "</Body>"
            dRng.InsertBefore "<Body originalStyle=""dowcipy"">"
        End If
    Next dPar
    Set dPar = Nothing
    Set dRng = Nothing
End Sub


Sub bibliografia_xml()
Dim iPar As Paragraph
Dim iRng As Range
    For Each iPar In ActiveDocument.Range.Paragraphs
        Set iRng = iPar.Range
        iRng.End = iRng.End - 1
        If iRng.Font.Size = 9 _
           And iRng.Font.Italic = True _
           And iRng.ParagraphFormat.Alignment = wdAlignParagraphRight _
           And Len(iRng) > 1 _
            Then
            iRng.InsertAfter "</Body>"
            iRng.InsertBefore "<Body originalStyle=""bibliografia"">"
        End If
    Next iPar
    Set iPar = Nothing
    Set iRng = Nothing
End Sub


Sub tabela_komplet_coipoile()


'Sub add1_before_xml()
Dim bRng As Range
Dim oPar As Paragraph
  Set bRng = ActiveDocument.Range
  With bRng.Find
    With .Font
      .Name = "Arial Narrow"
      '.Size = 8
      .Underline = wdUnderlineSingle = False
    bRng.ParagraphFormat.Alignment = wdAlignParagraphLeft = True
    End With
    
    
    

    Do While .Execute(MatchWholeWord:=True)
      bRng.End = bRng.Paragraphs(1).Range.End - 1
      With bRng
        .Select
        .InsertBefore "<tr><td>"
        .InsertAfter "</td></tr>"
        .Collapse 0
      End With
    Loop
  End With
  For Each oPar In ActiveDocument.Range.Paragraphs
    If oPar.Range.Characters.First = "-" And Not InStr(oPar.Range.Text, "<tr>") > 0 _
    And oPar.Range.Font.Size = 8 Then
      oPar.Range.InsertBefore "<tr><td>"
      oPar.Range.InsertAfter "</td></tr>"
    End If
  Next
lblb_Exit:
  Set bRng = Nothing
  
  'tabela_tabulator
  
   Dim aRng As Range
    Set aRng = ActiveDocument.Range
    With aRng.Find
        
                
                ' With .Font
                    '.Name = "Arial Narrow"
                    '.Size = 8
                    '.bold = False
                    '.Italic = False
                  '  .Underline = wdUnderlineSingle = False
               ' End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                         
             With aRng.Find
                .Text = "^t"
                .Replacement.Text = "</td><td>"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
         
        Loop
    End With
lbla_Exit:
    Set aRng = Nothing


 Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = _
            "</td><td>"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll


'Sub add_sign_xml22()
Dim sPar As Paragraph
    For Each sPar In ActiveDocument.Range.Paragraphs
        If Not sPar.Range = ActiveDocument.Range.Paragraphs.last.Range Then
            If sPar.Range.Characters.First = "<" _
               And sPar.Range.Underline = wdUnderlineSingle _
               And Not InStrRev(sPar.Range.Text, "-") = 1 _
               And Not InStr(sPar.Next(1).Range.Text, "<tr>") > 0 Then
                
                
                sPar.Range.InsertAfter "</tbody></table>]]></Body>"
            End If
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
    
    
    'Sub bodybodybold()

   
Dim lPar As Paragraph
    For Each lPar In ActiveDocument.Range.Paragraphs
       
            If lPar.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify = Center _
               And lPar.Range.Font.Name = "Georgia" _
               And Len(lPar.Range) > 2 Then
                lPar.Range.InsertBefore "<Subtitle>"
                lPar.Range.Characters.last.InsertBefore "</Subtitle>"
              
            
               
            End If
       
    Next
    
    
    'Sub dodaj_znacznik_tabeli_xml()
Dim zRng As Range
'Dim zPar As Paragraph
  Set zRng = ActiveDocument.Range
  With zRng.Find
      With .Font
      '.Size = 8
      .Underline = wdUnderlineSingle = False
       End With
       With zRng.Find
       .Text = "<tr><td>tucznik kl I</td>"
       End With
 
    
    Do While .Execute(MatchWholeWord:=True)
      zRng.End = zRng.Paragraphs(1).Range.End - 1
      With zRng
        .Select
        .InsertBefore "<Body><![CDATA[<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:300px;""><tbody>"
        .Collapse 0
      End With
    Loop
  End With
  
lblz_Exit:
  Set zRng = Nothing
  
  
' begin_table1
'
 Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<tr><td>Ceny paliw podajemy"
        .Replacement.Text = _
            "</tbody></table>]]></Body><Body>Ceny paliw podajemy"
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Lewicka - Jastrz¹b</td></tr>"
        .Replacement.Text = _
            "Lewicka - Jastrz¹b</Body>"
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<tr><td>Dane z"
        .Replacement.Text = _
            "<Body><![CDATA[<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:500px;""><tbody>^&^p"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Lewicka-Jastrz¹b</td></tr>"
        .Replacement.Text = _
            "Lewicka-Jastrz¹b</Body>"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Subtitle>Ceny ¿ywca"
        .Replacement.Text = _
            "</tbody></table>]]></Body>^&^p"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Body originalStyle=""body_ramkaj"">Ceny skupu"
        .Replacement.Text = _
            "</tbody></table>]]></Body>^&"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Ceny skupu ¿ywca wieprzowego"
        .Replacement.Text = _
            "</tbody></table>]]></Body>^p <Body><![CDATA[<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:300px;""><tbody>^&"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<tr><td>Ceny paliw podajemy z wybieranych co tydzieñ stacji paliw z  terenu powiatu ropczycko – sêdziszowskiego. Opracowanie Sylwia Lewicka-Jastrz¹b</td></tr>"
        .Replacement.Text = _
            "Ceny paliw podajemy z wybieranych co tydzieñ stacji paliw z  terenu powiatu ropczycko - sêdziszowskiego. Opracowanie Sylwia Lewicka-Jastrz¹b"
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Ceny skupu ¿ywca w Kabanospol Tel. 17 22 14 800"
        .Replacement.Text = _
            "<Body>Ceny skupu ¿ywca w Kabanospol Tel. 17 22 14 800</Body>"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<tr><td>Podawane ceny obowi¹zuj¹"
        .Replacement.Text = _
            "<Body>Podawane ceny obowi¹zuj¹"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ród³o: informacja w³asna</td></tr>"
        .Replacement.Text = _
            "informacja w³asna</Body>"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Body>Podawane ceny obowi¹zuj¹ w dniu"
        .Replacement.Text = _
            "</tbody></table>]]></Body>^&^p"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    
    
    
  
  
  
  
  
  
  
  'Sub add_xml_intro_outro1()


ActiveDocument.Content.InsertBefore "<Article>" & Chr(13) & "<Story>" & Chr(13) & "<Title>Co i po ile</Title>" & Chr(13)
ActiveDocument.Content.InsertAfter "</Story>" & Chr(13) & "</Article>"

 Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Title>Co i po ile</Title>"
        .Replacement.Text = _
            "^&<Body><![CDATA[<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:500px;""><tbody>"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll


End Sub

Sub oglo_modul()

   
Dim cPar As Paragraph
    For Each cPar In ActiveDocument.Range.Paragraphs
       
            If cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter _
               And cPar.Range.Font.Size = 10 _
               And Len(cPar.Range) > 2 _
               And cPar.Range.Font.Name = "Arial Narrow" Then
                cPar.Range.InsertBefore "<Body originalStyle=""oglo_modul"">"
                cPar.Range.Characters.last.InsertBefore "</Body><Body><![CDATA[<br />]]></Body>"
              
            
               
            End If
       
    Next
        End Sub

Sub oglo_naglowki()



    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "RÓ¯NE: ^p^tSPRZEDAM^p^tODDAM"
        .Replacement.Text = "RÓ¯NE: SPRZEDAM, ODDAM"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "NIERUCHOMOŒCI^p^tSPRZEDAM^p^tMIESZKANIE:"
        .Replacement.Text = "NIERUCHOMOŒCI: SPRZEDAM MIESZKANIE:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "NIERUCHOMOŒCI^p^tKUPIÊ"
        .Replacement.Text = "NIERUCHOMOŒCI KUPIÊ:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "NIERUCHOMOŒCI^p^tWYNAJMÊ"
        .Replacement.Text = "NIERUCHOMOŒCI WYNAJMÊ:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "US£UGI^pOGÓLNOBUDOWLANE:"
        .Replacement.Text = "US£UGI OGÓLNOBUDOWLANE:"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "US£UGI^p^tRÓ¯NE"
        .Replacement.Text = "US£UGI RÓ¯NE"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "PRACA^p^tDAM"
        .Replacement.Text = "PRACA DAM"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "PRACA^p^tSZUKAM"
        .Replacement.Text = "PRACA SZUKAM"
        .Forward = True
        .Wrap = wdFindContinue
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    End Sub
    
    
    Sub oglo_naglowki_xml()

    
Dim hPar As Paragraph
Dim hRng As Range
    For Each hPar In ActiveDocument.Range.Paragraphs
        Set hRng = hPar.Range
        hRng.End = hRng.End - 1
        If Len(hRng) > 1 _
           And hRng.Font.Color = RGB(33, 33, 32) = False Then
            hRng.InsertAfter "</Subtitle>"
            hRng.InsertBefore "<Subtitle originalStyle=""oglo_tyt"">"
        End If
    Next hPar
    Set hPar = Nothing
    Set hRng = Nothing
    
    End Sub
    
    

Sub Makro11()
'
' Makro11 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Title>Co i po ile</Title>"
        .Replacement.Text = _
            "^&<Body><![CDATA[<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:400px;""><tbody>"
        .Forward = True
        .Wrap = wdFindContinue
           End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub zdjecie_oglo_xml()

Dim check As Boolean
Dim search As String
Dim para As Paragraph
Dim tempStr As String
Dim txt As String

search = "zdjêcie"

For Each para In ActiveDocument.Paragraphs
    txt = para.Range.Text
    tempStr = (txt)
    check = InStr(tempStr, search)

    If check = True Then
        If Not para.Range = ActiveDocument.Range.Paragraphs.First.Range Then
              If Len(para.Range) > 2 _
               And para.Range.Font.Name = "Arial Narrow" Then
                para.Previous(1).Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)

                'para.Previous(1).Range.Characters.last.InsertBefore "</Body>"
                End If
                End If
    End If
Next
End Sub


Sub srodtytul()

   
Dim cPar As Paragraph
    For Each cPar In ActiveDocument.Range.Paragraphs
       
            If cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify _
               And cPar.Range.Font.Size = 9 _
                                And cPar.Range.Font.Name = "Georgia" _
                 And Len(cPar.Range) > 10 _
                 And Len(cPar.Range) < 60 _
                 And Not cPar.Range.Characters.last = "." Then
             
                cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                cPar.Range.InsertBefore Chr(13)
                cPar.Range.Font.Size = 10
                cPar.Range.Font.Bold = True
              
            
                'And cPar.Range.Font.bold = True
            End If
       
    Next
        End Sub


Sub oglo_modul_test_srodtytul()

   
Dim cPar As Paragraph
    For Each cPar In ActiveDocument.Range.Paragraphs
       
            If cPar.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify _
               And cPar.Range.Font.Size = 10 _
               And Not Len(cPar.Range) < 5 _
               And cPar.Range.Font.Bold = True _
               And cPar.Range.Font.Name = "Georgia" Then
                cPar.Range.InsertBefore "<Body originalStyle=""oglo_modul"">"
                cPar.Range.Characters.last.InsertBefore "</Body><Body><![CDATA[<br />]]></Body>"
              
            
               
            End If
       
    Next
        End Sub


Sub Makro12()
'
' Makro12 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>DOWCIPY" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "<Title>DOWCIPY" & ChrW(9660) & "</Title>"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>DOWCIPY" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "<Title>DOWCIPY" & ChrW(9660) & "</Title>"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
End Sub
Sub Makro13()
'
' Makro13 Makro
'
'
    
    Selection.Range.Hyperlinks(1).Delete
End Sub
Sub Makro14()
'
' Makro14 Makro
'
'
    Selection.WholeStory
    Selection.Fields.Unlink
End Sub

Sub zdjecie_oglo_xml_opinia()

Dim check As String
Dim search As Variant
Dim para As Paragraph
Dim tempStr As String
Dim txt As String

search = Array("OPINIA", "OPINIE")

For Each para In ActiveDocument.Paragraphs
    txt = para.Range.Text
    tempStr = (txt)
    check = InStr(tempStr, search)

    If check = True Then
            If Len(para.Range) > 2 _
               And para.Range.Font.Name = "Times New Roman" _
               And para.Range.Font.Italic = False Then
                para.Next(4).Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)
             End If
        
    End If
Next
End Sub


Sub zdjecie_oglo_xml_opinia_array()

Dim check As Boolean
Dim search() As Variant
Dim para As Paragraph
Dim tempStr As String
Dim txt As String

search = Array("OPINIA", "OPINIE")

For Each para In ActiveDocument.Paragraphs
    txt = para.Range.Text
    tempStr = (txt)
    check = InStr(tempStr, search)

    If check = True Then
        If Not para.Range = ActiveDocument.Range.Paragraphs.First.Range Then
              If Len(para.Range) > 2 _
               And para.Range.Font.Name = "Times New Roman" _
               And para.Range.Font.Italic = False Then
                para.Next(4).Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)

               
                End If
                End If
    End If
Next
End Sub


Sub zdjecie_oglo_xml_opinia_rrr()

Dim check As Boolean
Dim search As String
Dim para As Paragraph
Dim tempStr As String
Dim txt As String

search = "OPINIA"

For Each para In ActiveDocument.Paragraphs
    txt = para.Range.Text
    tempStr = (txt)
    check = InStr(tempStr, search)

    If check = True Then
        
              If Len(para.Range) > 2 _
               And para.Range.Font.Name = "Times New Roman" _
               And para.Range.Font.Italic = False Then
                para.Next(4).Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)

                'para.Previous(1).Range.Characters.last.InsertBefore "</Body>"
                End If
                
    End If
Next
End Sub

Sub zdjecie_oglo_xml_opinia_rrr11()



Dim parao As Paragraph
For Each parao In ActiveDocument.Paragraphs
    
    If InStr(parao, "OPINIA") > 0 Or InStr(parao, "OPINIE") > 0 Then
        
              If Len(parao.Range) > 2 _
               And parao.Range.Font.Name = "Times New Roman" _
               And parao.Range.Font.Italic = False Then
                parao.Next(3).Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)
              End If
                
    End If
Next

End Sub

Sub pogrubianie_nazw_miejscowosci()


      'Application.ScreenUpdating = False
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("ROPCZYCE", "IWIERZYCE", "OSTRÓW", "BÊDZIENICA", "BYSTRZYCA", "NOCKOWA", _
    "OLCHOWA", "OLIMPÓW", "SIELEC", "WIERCANY", "WIŒNIOWA", "BRZEZÓWKA", "GNOJNICA", "LUBZINA", _
    "MA£A", "NIEDWIADA", "OKONIN", "BLIZNA", "KAMIONKA", "KOZODRZA", "OCIEKA", "SKRZYSZÓW", _
    "ZD¯ARY", "BÊDZIEMYŒL", "BORECZEK", "BUKOWINA", "CIERPISZ", "KAWÊCZYN", _
     "KLÊCZANY", "KRZYWA", "RUDA", "SZKODNA", _
    "ZAB£OCIE", "ZAGORZYCE", "BRONISZÓW", "BRZEZINY", "GLINIK", "NAWSIE", "RZESZÓW", "PASZCZYNA", _
    "WIELOPOLE SKRZYÑSKIE", "SÊDZISZÓW M£P.", "BOREK WIELKI", "BOREK MA£Y", "WARSZAWA", _
"£¥CZKI KUCHARSKIE", "CZARNA SÊDZISZOWSKA", "WOLICA £UGOWA", "WOLICA PIASKOWA", _
"GÓRA ROPCZYCKA", "KAWÊCZYN SÊDZISZOWSKI", "WOLA OCIECKA", "CA£Y POWIAT", "SÊDZISZÓW MA£OPOLSKI")
    
                  
                
    
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
                .Format = True
                .MatchWildcards = True
                .Font.Bold = True
                .Font.Name = "Georgia"
                .Font.Size = 9
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

End Sub

Sub italic_xml()


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
    
    End Sub





Public Sub FindReplaceAnywhere()

  Dim rngStory As Word.Range
  Dim pFindTxt As String
  Dim pReplaceTxt As String
  Dim lngJunk As Long
  Dim oShp As Shape
  Dim par As Paragraph
  pFindTxt = InputBox("Wpisz szukan¹ nazwê zdjêcia" _
    , "SZUKAJ")
  If pFindTxt = "" Then
    MsgBox "Cancelled by User"
    Exit Sub
  End If
TryAgain:
  pReplaceTxt = InputBox("Wpisz zastêpowan¹ nazwê", "ZAMIEÑ")
  If pReplaceTxt = "" Then
    If MsgBox("Do you just want to delete the found text?", _
     vbYesNoCancel) = vbNo Then
      GoTo TryAgain
    ElseIf vbCancel Then
      MsgBox "Cancelled by User."
      Exit Sub
    End If
  End If
  'Fix the skipped blank Header/Footer problem
  lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType
  'Iterate through all story types in the current document
  For Each rngStory In ActiveDocument.StoryRanges
    'Iterate through all linked stories
    Do
      SearchAndReplaceInStory rngStory, pFindTxt, pReplaceTxt
      On Error Resume Next
      Select Case rngStory.StoryType
      Case 6, 7, 8, 9, 10, 11
        If rngStory.ShapeRange.Count > 0 Then
          For Each oShp In rngStory.ShapeRange
            If oShp.TextFrame.HasText Then
              SearchAndReplaceInStory oShp.TextFrame.TextRange, _
                  pFindTxt, pReplaceTxt
            End If
          Next
        End If
      Case Else
        'Do Nothing
      End Select
      On Error GoTo 0
      'Get next linked story (if any)
      Set rngStory = rngStory.NextStoryRange
    Loop Until rngStory Is Nothing
  Next
End Sub

Public Sub SearchAndReplaceInStory(ByVal rngStory As Word.Range, _
    ByVal strSearch As String, ByVal strReplace As String)
  With rngStory.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = strSearch
    .Replacement.Text = strReplace
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceOne
  End With
End Sub



Sub highlight_xml()


Dim iPar As Paragraph
    For Each iPar In ActiveDocument.Range.Paragraphs
       
            If Not iPar.Range.Characters.First = "<" _
            And Not iPar.Range.Characters.last = ">" _
            Then
                iPar.Range.HighlightColorIndex = wdBrightGreen
              
            
                       End If
       
    Next
    
    End Sub
    
    Sub podpisu_xml()

 

With Selection.Find
        .Text = "CA£KIEM (NIE) OBIEKTYWNIE Wojciech Naja"
        .Wrap = wdFindContinue
          End With
            
    With Selection.Find
   Selection.ClearFormatting
   .Replacement.Text = "^&" & Chr(13) & "<Author>Wojciech Naja</Author>^p"
   .Execute Replace:=wdReplaceAll
    End With
        
   End Sub
Sub Makro15()
'
' Makro15 Makro
'
'
    Selection.ClearFormatting
    WordBasic.ClearallFormatting
End Sub


Sub ClrFmtgReplace()
 Dim rngTemp As Range
 Set rngTemp = ActiveDocument.Content
 With rngTemp.Find
 .ClearFormatting
 .Replacement.ClearFormatting
 .MatchWholeWord = True
 .Execute findText:="CA£KIEM (NIE) OBIEKTYWNIE Wojciech Naja", ReplaceWith:="^&" & Chr(13) & "<Author>Wojciech Naja</Author>^p", _
 Replace:=wdReplaceAll
 .ClearFormatting
 End With
End Sub

Sub Makro16()
'
' Makro16 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>REPORTER I PIENI¥DZE" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "^&^p<Author>Pawe³ Bochenek</Author> "
        .Forward = True
        
         End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub
Sub Makro17()
'
' autorzy_niepodpisani
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>REPORTER I PIENI¥DZE" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "^&^p<Author>Pawe³ Bochenek</Author> "
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
End Sub
Sub Makro18()
'
' Makro18 Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 10
        .Bold = False
        .Italic = False
        .Name = "Georgia"
    End With
    With Selection.Find.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = "<Subtitle>^&</Subtitle>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub przepis_xml()

    
Dim sPar As Paragraph
Dim sRng As Range
    For Each sPar In ActiveDocument.Range.Paragraphs
        Set sRng = sPar.Range
        sRng.End = sRng.End - 1
        If Len(sRng) > 10 _
           And sRng.Font.Name = "Georgia" _
           And sRng.Font.Size = 10 _
            And Not sRng.Font.Bold = True _
           And sRng.ParagraphFormat.Alignment = wdAlignParagraphCenter Then
            'sPar.Range.HighlightColorIndex = wdDarkRed
            sRng.InsertAfter "</Subtitle>"
           sRng.InsertBefore "<Subtitle>"
        End If
    Next sPar
    Set sPar = Nothing
    Set sRng = Nothing
    
    End Sub
    
    Sub body_xml()

    

    
Dim sPar As Paragraph
Dim sRng As Range
    For Each sPar In ActiveDocument.Range.Paragraphs
        Set sRng = sPar.Range
        sRng.End = sRng.End - 1
        If Len(sRng) > 10 _
           And sRng.Font.Name = "Georgia" _
           And sRng.Font.Size = 9 _
            And Not sRng.Font.Bold = True _
            And Not sRng.Font.Italic = True _
           And sRng.ParagraphFormat.Alignment = wdAlignParagraphJustify _
           And Not sRng.Characters.First = "<" Then
            'sPar.Range.HighlightColorIndex = wdDarkRed
            sRng.InsertAfter "</Body>"
           sRng.InsertBefore "<Body>"
        End If
    Next sPar
    Set sPar = Nothing
    Set sRng = Nothing
    
    End Sub
    
    
    Sub zdjecie_oglo_xml_opinia_za_kazdym_razem()



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
                parax.Next.Range.InsertAfter "<Picture><Image href=""file://images/nazwa_zdjecia.jpg""></Image></Picture>" & Chr(13) _
                & "<Subtitle originalStyle=""linia_odpychanie"">x</Subtitle>" & Chr(13)
              End If
                
    End If
Next

End Sub
 
 
Sub zajawka_publico()





Dim iPar As Paragraph
    For Each iPar In ActiveDocument.Range.Paragraphs
       
            If iPar.Range.Font.Bold = True _
               And Len(iPar.Range) > 1 _
               Then
                iPar.Range.InsertBefore "<b>"
                iPar.Range.Characters.last.InsertBefore "</b><br />"
              
            
            End If
           
       
    Next
    
    Dim bPar As Paragraph
    For Each bPar In ActiveDocument.Range.Paragraphs
       
            If bPar.Range.Font.Bold = False _
               And Len(bPar.Range) > 1 _
               And Not bPar.Range = ActiveDocument.Range.Paragraphs.last.Range Then
              
               bPar.Range.Characters.last.InsertBefore "<br /><br />"
              
            
            End If
           
       
    Next

    
    
    End Sub
Sub Findfirstcharacterinpara()
Dim wdoc As Document
Dim para As Paragraph
Set wdoc = ActiveDocument
For Each para In wdoc.Paragraphs
If para.Range.Characters(1) = Chr(32) Then para.Range.Characters(1).Delete
Next para
End Sub

Sub body_awaria()


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

End Sub
Sub ceny_paliw()
'
' ceny_paliw Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
    
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
    
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
    
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
    
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
    
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
     
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub Makro19()
'
' Makro19 Makro
'
'
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.1)
        .RightIndent = CentimetersToPoints(0.1)
        .SpaceBefore = 6
        .SpaceAfter = 6
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .TextboxTightWrap = wdTightNone
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.83) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
End Sub

Sub RAMkA_oglo()



Dim check As Boolean
Dim search As String
Dim para As Paragraph
Dim tempStr As String
Dim txt As String

search = "RAMKA"

For Each para In ActiveDocument.Paragraphs
    txt = para.Range.Text
    tempStr = (txt)
    check = InStr(tempStr, search)

    If check = True Then
        With para.Range
        With .ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.1)
        .RightIndent = CentimetersToPoints(0.1)
        .SpaceBefore = 6
        .SpaceAfter = 6
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .TextboxTightWrap = wdTightNone
    End With
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3.83) _
        , Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    End If
Next
End Sub


Sub body_loop_xml_error()


          
    Dim bRng As Range
    Set bRng = ActiveDocument.Range
    With bRng.Find
        
                 With .ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                End With
                 With .Font
                    .Size = 9
                    .Bold = False
                End With
          
        Do While .Execute(MatchWholeWord:=True)
       
                
         bRng.End = bRng.Paragraphs(1).Range.End - 1
             With bRng
                .InsertAfter "<br /><br />]]></Body>"
                .Collapse 0
            End With
         
        Loop
    End With
lblb_Exit:
    Set bRng = Nothing
   End Sub


Sub body_loop_rozmowa_xml()


 'wstawianie do znaczników kiedy nie ma pe³nego formatowania paragrafu
   
               
Dim sPara As Paragraph
Dim sRnga As Range
    For Each sPara In ActiveDocument.Range.Paragraphs
        Set sRnga = sPara.Range
        sRnga.End = sRnga.End - 1
        If Len(sRnga) > 1 _
            And sRnga.Font.Size = 9 _
            And sRnga.ParagraphFormat.Alignment = wdAlignParagraphLeft _
            And sRnga.Words.First.Bold = True _
            And Not sRnga.Characters.First = "<" _
           And Not sRnga.Characters.last = ">" Then
            
            sRnga.InsertAfter "</Body>"
         
        End If
    Next sPara
    Set sPara = Nothing
    Set sRnga = Nothing

End Sub


Sub cogdziekiedy_xml()


  'Sub add_after_bold8888()

 Dim oRng As Range
    Set oRng = ActiveDocument.Range
    With oRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Font.Bold = True
        .Font.Color = RGB(33, 33, 32)
        Do While .Execute
        If Len(oRng) > 3 Then
            oRng.Text = "{body:bold}" & oRng.Text & "{/body:bold}"
            oRng.Collapse wdCollapseEnd
            End If
        Loop
    End With


 
    
    'czerwony kolor
   
               
Dim sPara As Paragraph
Dim sRnga As Range
    For Each sPara In ActiveDocument.Range.Paragraphs
        Set sRnga = sPara.Range
        sRnga.End = sRnga.End - 1
        If Len(sRnga) > 1 _
            And sRnga.ParagraphFormat.Alignment = wdAlignParagraphCenter _
            And sRnga.Words.First.Bold = True _
            And Not sRnga.Characters.First = "<" _
           And Not sRnga.Characters.last = ">" _
           And sRnga.Font.Color = RGB(231, 72, 45) Then
            
            sRnga.InsertAfter "</Body>"
            sRnga.InsertBefore "<Body originalStyle=""red_arial"">"
         
        End If
    Next sPara
    Set sPara = Nothing
    Set sRnga = Nothing
    
    'Sub cogdziekiedy_xml2()


 'niebieski
   
               
Dim nPara As Paragraph
Dim nRnga As Range
    For Each nPara In ActiveDocument.Range.Paragraphs
        Set nRnga = nPara.Range
        nRnga.End = nRnga.End - 1
        If Len(nRnga) > 1 _
            And nRnga.ParagraphFormat.Alignment = wdAlignParagraphJustify _
                        And Not nRnga.Characters.First = "<" _
           And Not nRnga.Characters.last = ">" _
           And nRnga.Font.Color = RGB(79, 186, 226) Then
            
            nRnga.InsertAfter "</Body>"
            nRnga.InsertBefore "<Body originalStyle=""bule_light_arial"">"
         
        End If
    Next nPara
    Set nPara = Nothing
    Set nRnga = Nothing
    
    'Sub cogdziekiedy_xml23()


 'czarny
   
               
Dim yPara As Paragraph
Dim yRnga As Range
    For Each yPara In ActiveDocument.Range.Paragraphs
        Set yRnga = yPara.Range
        yRnga.End = yRnga.End - 1
        If Len(yRnga) > 1 _
            And yRnga.ParagraphFormat.Alignment = wdAlignParagraphJustify _
                        And Not yRnga.Characters.First = "<" _
           And Not yRnga.Characters.last = ">" _
           And yRnga.Font.Bold = True Then
            
            yRnga.InsertAfter "</Body>"
            yRnga.InsertBefore "<Body originalStyle=""bold"">"
         
        End If
    Next yPara
    Set yPara = Nothing
    Set yRnga = Nothing
    
    
    
'Sub cogdziekiedy_xml234()


 'czarny
   
               
Dim cPara As Paragraph
Dim cRnga As Range
    For Each cPara In ActiveDocument.Range.Paragraphs
        Set cRnga = cPara.Range
        cRnga.End = cRnga.End - 1
        If Len(cRnga) > 1 _
            And cRnga.ParagraphFormat.Alignment = wdAlignParagraphCenter _
            And cRnga.Underline = wdUnderlineSingle _
                        And Not cRnga.Characters.First = "<" _
           And Not cRnga.Characters.last = ">" _
           And cRnga.Words.First.Bold = True _
            And cRnga.Font.Color = RGB(33, 33, 32) Then
            cRnga.InsertAfter "</Body>"
            cRnga.InsertBefore "<Body originalStyle=""body_ramkaj"">"
         
        End If
    Next cPara
    Set cPara = Nothing
    Set cRnga = Nothing
    
    'Sub cogdziekiedy_xml2345()


 'czarny
   
               
Dim iPara As Paragraph
Dim iRnga As Range
    For Each iPara In ActiveDocument.Range.Paragraphs
        Set iRnga = iPara.Range
        iRnga.End = iRnga.End - 1
        If Len(iRnga) > 1 _
           And Not iRnga.Characters.First = "<" _
           And Not iRnga.Characters.last = ">" Then
          
            iRnga.InsertAfter "</Body>"
            iRnga.InsertBefore "<Body>"
         
        End If
    Next iPara
    Set iPara = Nothing
    Set iRnga = Nothing

End Sub


Sub cogdziekiedy_xml2()


 'niebieski
   
               
Dim nPara As Paragraph
Dim nRnga As Range
    For Each nPara In ActiveDocument.Range.Paragraphs
        Set nRnga = nPara.Range
        nRnga.End = nRnga.End - 1
        If Len(nRnga) > 1 _
            And nRnga.ParagraphFormat.Alignment = wdAlignParagraphJustify _
                        And Not nRnga.Characters.First = "<" _
           And Not nRnga.Characters.last = ">" _
           And nRnga.Font.Color = RGB(79, 186, 226) Then
            
            nRnga.InsertAfter "</Body>"
            nRnga.InsertBefore "<Body originalStyle=""bule_light_arial"">"
         
        End If
    Next nPara
    Set nPara = Nothing
    Set nRnga = Nothing

End Sub




Sub cogdziekiedy_xml23()


 'czarny
   
               
Dim nPara As Paragraph
Dim nRnga As Range
    For Each nPara In ActiveDocument.Range.Paragraphs
        Set nRnga = nPara.Range
        nRnga.End = nRnga.End - 1
        If Len(nRnga) > 1 _
            And nRnga.ParagraphFormat.Alignment = wdAlignParagraphJustify _
                        And Not nRnga.Characters.First = "<" _
           And Not nRnga.Characters.last = ">" _
           And nRnga.Font.Bold = True Then
            
            nRnga.InsertAfter "</Body>"
            nRnga.InsertBefore "<Body originalStyle=""bold"">"
         
        End If
    Next nPara
    Set nPara = Nothing
    Set nRnga = Nothing

End Sub


Sub cogdziekiedy_xml234()


 'czarny
   
               
Dim cPara As Paragraph
Dim cRnga As Range
    For Each cPara In ActiveDocument.Range.Paragraphs
        Set cRnga = cPara.Range
        cRnga.End = cRnga.End - 1
        If Len(cRnga) > 1 _
            And cRnga.ParagraphFormat.Alignment = wdAlignParagraphCenter _
            And cRnga.Underline = wdUnderlineSingle _
                        And Not cRnga.Characters.First = "<" _
           And Not cRnga.Characters.last = ">" _
           And cRnga.Words.First.Bold = True _
            And cRnga.Font.Color = RGB(33, 33, 32) Then
            cRnga.InsertAfter "</Body>"
            cRnga.InsertBefore "<Body originalStyle=""body_ramkaj"">"
         
        End If
    Next cPara
    Set cPara = Nothing
    Set cRnga = Nothing

End Sub

Sub cogdziekiedy_xml2345()


 'czarny
   
               
Dim iPara As Paragraph
Dim iRnga As Range
    For Each iPara In ActiveDocument.Range.Paragraphs
        Set iRnga = iPara.Range
        iRnga.End = iRnga.End - 1
        If Len(iRnga) > 1 _
           And Not iRnga.Characters.First = "<" _
           And Not iRnga.Characters.last = ">" Then
          
            iRnga.InsertAfter "</Body>"
            iRnga.InsertBefore "<Body>"
         
        End If
    Next iPara
    Set iPara = Nothing
    Set iRnga = Nothing

End Sub

 Sub add_after_bold8888()

 


  
    Dim x As Long, i As Long, ArrFnd()
    ArrFnd = Array("")
    For x = 0 To UBound(ArrFnd)
        With ActiveDocument.Range
            With .Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = True
                .Font.Bold = True
                .Font.Color = RGB(33, 33, 32)
                .ParagraphFormat.Alignment = wdAlignParagraphJustify
                .Execute
            End With
            Do While .Find.Found
                .InsertAfter "{/body:bold}"
                .InsertBefore "{body:bold}"
                .End = .End
                .Collapse wdCollapseEnd
                .Find.Execute
            Loop
        End With
    Next
  
    
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "{body:bold}^p{/body:bold}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub boldblolod()

 'czarny
   
               
Dim iRnga As Range
    For Each iRnga In ActiveDocument.Range
       Set iRnga = iRng.Range
        'iRnga.End = iRnga.End - 1
        If Len(iRnga) > 2 _
        And iRnga.Bold = True Then
                      iRnga.InsertAfter "{/body:bold}"
            iRnga.InsertBefore "{body:bold}"
         
        End If
    Next iRnga
    
    Set iRnga = Nothing

End Sub



Sub Makro20()
'
' Makro20 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^s^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub add_bold2()
Dim oRng As Range
    Set oRng = ActiveDocument.Range
    With oRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Font.Bold = True
        .Font.AllCaps = True
       '.Font.Color = RGB(33, 33, 32)
       
        Do While .Execute
        If Len(oRng) > 3 Then
            oRng.Text = "{body:bold}" & oRng.Text & "{/body:bold}"
            oRng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub


Sub red_test()
Dim oRng As Range
    Set oRng = ActiveDocument.Range
    With oRng.Find
        '.ClearFormatting
        '.Replacement.ClearFormatting
        .Text = ""
        '.Font.bold = True
        '.Font.AllCaps = True
        '.Font.SmallCaps = False
         
         Do While .Execute
        If UCase(oRng) Then
            oRng.Text = "{body:bold}" & oRng.Text & "{/body:bold}"
            oRng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub Makro21()
'
' Makro21 Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = True
        .AllCaps = False
        .Superscript = False
        .Subscript = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
Sub Makro22()
'
' Makro22 Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub

  
    
  Sub FOTO_name_loop_xml()
          
    
    'Sub JPG_xml()

    
             
    
    
    
    'Sub JPG_xml()

    
             
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
                .InsertAfter "</Description>^p</Picture>"
                .Collapse 0
            End With
         
        Loop
    End With
lbld_Exit:
    Set dRng = Nothing
   'Exit Sub


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
    Set jRng = Nothing
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
    
    End Sub
    
    
    Sub JPG_xml()

    
             
    Dim fRng As Range
    
    Set fRng = ActiveDocument.Range
    With fRng.Find
         
        Do While .Execute(findText:="^13<*>.JPG", _
        MatchWholeWord:=False, MatchWildcards:=True, _
        MatchCase:=False)
     
   'fRng.End = fRng.End - 1
   'fRng.Start = fRng.Start + 2
            With fRng
                            
              .Start = fRng.Start + 1
                .InsertAfter """></Image>" & Chr(13)
                
                .InsertBefore "<Picture>" & Chr(13) & "<Image href=""file://images/"
                                  
              
                .Collapse 0
            End With
            
    
            'Exit Do
        Loop
    End With
    
lblx_Exit:
    Set fRng = Nothing
    End Sub
    
    Sub test_foto_jpg_xml()
Dim ArrFind As Variant
Dim i As Long

ArrFind = Array("^13<*>.JPG", "^13<*>.jpg")

For i = 0 To UBound(ArrFind)
Selection.HomeKey wdStory
With Selection.Find
 .ClearFormatting
    Do While .Execute(findText:=ArrFind(i), MatchWholeWord:=False, MatchWildcards:=True)
     
   'fRng.End = fRng.End - 1
   'fRng.Start = fRng.Start + 2
            With fRng
                            
                .Start = fRng.Start + 1
                .InsertAfter """></Image>" & Chr(13)
                
                .InsertBefore "<Picture>" & Chr(13) & "<Image href=""file://images/"
                                  
              
                .Collapse 0
            End With
            
    
            'Exit Do
        Loop
    End With
   Next i
lblx_Exit:
    Set fRng = Nothing
    
End Sub

Sub Makro23()
'
' Makro23 Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^13<*>.JPG"
        .Replacement.Text = "<body>^&</body>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub JPGbbb_xml_1()
    Dim fRng        As Range
    
    Set fRng = ActiveDocument.Range
    
        Do While fRng.Find.Execute(findText:=".JPG", _
                          MatchWholeWord:=True, MatchCase:=False)


            Set fRng = fRng.Paragraphs.Item(1).Range
            
            'fRng.Select 'tylko na czas testów!!!
            
            With fRng
                .Text = "<Picture>" & Chr(13) & "<Image href=""file://images/" & Left(.Text, Len(.Text) - 1) & """></Image>" & String(2, Chr(13))
                .Collapse 0
            End With


        Loop


lblx_Exit:
    Set fRng = Nothing
End Sub

Sub Publico_xml_sport_foto_filename()

 ' usuwanie spacji bia³ych
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
    .Text = "Miko³aj Froñ"
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
        .Text = "Marcin Jastrzêbski"
        .Replacement.Text = "^p<Author>Marcin Jastrzêbski</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^tso"
        .Replacement.Text = "^p<Author>S³awomir Oskarbski</Author>^p"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "^tdc"
        .Replacement.Text = "^p<Author>Dominika Czy¿</Author>^p"
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
        .Replacement.Text = "^p<Author>Miko³aj Froñ</Author>^p"
        .Wrap = wdFindContinue
         End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
       
        .Text = "^tjas"
        .Replacement.Text = "^p<Author>Marcin Jastrzêbski</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
     With Selection.Find
    .Text = "^ttab"
        .Replacement.Text = "^p<Author>W³adys³aw Tabasz</Author>"
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
        .Replacement.Text = "^p<Author>Pawe³ Bochenek</Author>"
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
        .Replacement.Text = "^p<Author>Dominika Czy¿</Author>"
        .Wrap = wdFindContinue
          End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      
     With Selection.Find
    .Text = "Ortalion"
        .Replacement.Text = "^p<Author>Bronis³aw Róg</Author>"
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
    ArrFnd = Array("ROPCZYCE", "IWIERZYCE", "OSTRÓW", "BÊDZIENICA", "BYSTRZYCA", "NOCKOWA", _
    "OLCHOWA", "OLIMPÓW", "SIELEC", "WIERCANY", "WIŒNIOWA", "BRZEZÓWKA", "GNOJNICA", "LUBZINA", _
    "MA£A", "NIEDWIADA", "OKONIN", "BLIZNA", "KAMIONKA", "KOZODRZA", "OCIEKA", "SKRZYSZÓW", _
    "ZD¯ARY", "BÊDZIEMYŒL", "BORECZEK", "BUKOWINA", "CIERPISZ", "KAWÊCZYN", _
     "KLÊCZANY", "KRZYWA", "RUDA", "SZKODNA", "TARNÓW", _
    "ZAB£OCIE", "ZAGORZYCE", "BRONISZÓW", "BRZEZINY", "GLINIK", "NAWSIE", "RZESZÓW", "PASZCZYNA", _
    "WIELOPOLE SKRZYÑSKIE", "SÊDZISZÓW M£P.", "BOREK WIELKI", "BOREK MA£Y", "WARSZAWA", _
"£¥CZKI KUCHARSKIE", "CZARNA SÊDZISZOWSKA", "WOLICA £UGOWA", "WOLICA PIASKOWA", _
"GÓRA ROPCZYCKA", "KAWÊCZYN SÊDZISZOWSKI", "WOLA OCIECKA", "CA£Y POWIAT", "SÊDZISZÓW MA£OPOLSKI", _
"SIATKÓWKA", "SUMO", "TENIS STO£OWY", "BOKS", "SZACHY", "KARATE", "HALOWA PI£KA NO¯NA", _
    "PI£KA NO¯NA", "PODNOSZENIE CIÊ¯ARÓW", "PI£KARSKIE WIEŒCI", "ZAPASY", "KOLARSTWO")
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
        .Text = "Prosimy traktowaæ poni¿szy horoskop z przymru¿eniem oka, gdy¿ prawdopodobieñstwo, ¿e opisane sytuacje kiedykolwiek zaistniej¹ jest znikome i zale¿y jedynie od przypadku."
        .Replacement.Text = "<Body originalStyle=""dowcipy"">^&</Body><Body><![CDATA[<br />]]></Body>"
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Wyra¿ane przez Czytelników opinie nie s¹ stanowiskiem redakcji Reporter Gazety."
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
        .Text = "<Vignette>REPORTER I PIENI¥DZE" & ChrW(9660) & "</Vignette>"
        .Replacement.Text = "^&^p<Author>Pawe³ Bochenek</Author> "
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
        .Text = "<Vignette>(TROCHÊ) M£ODSZYM OKIEM" & ChrW(9658) & " Miko³aj Froñ </Vignette>"
        .Replacement.Text = "^&^p<Author>Miko³aj Froñ</Author> "
        .Forward = True
        .Wrap = wdFindContinue
      
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<Vignette>CA£KIEM (NIE) OBIEKTYWNIE" & ChrW(9658) & " Wojciech Naja </Vignette>"
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
Sub Makro24()
'
' Makro24 Makro
'
'
   
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ", audiodeskrypcja [AD]"
        .Replacement.Text = ""
        .Forward = True
        
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub pusl()
'
' pusl Makro
'
'

End Sub


Sub pogrubienia_do_znakow_tvp_i_inne222()
    Dim r As Range
    
    Set r = ActiveDocument.Range
    
    'r.Font.bold = True

    With r.Find
        .MatchWildcards = True
        
        .Text = "Big Brother *[0-9]"
        .Replacement.Text = "Big Brother"
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub
Sub form()
'
' form Makro
'
'

End Sub
Sub trwam_czyszczenie()
'
' trwam_czyszczenie Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^#:^#^#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "S³owo ¯ycia - rozwa¿anie Ewangelii dnia"
        .Replacement.Text = "S³owo ¯ycia"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Z wêdk¹ nad wodê w Polskê i Œwiat"
        .Replacement.Text = "Z wêdk¹ nad wodê"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Z wêdk¹ nad wodê w Polskê i Œwiat"
        .Replacement.Text = "Z wêdk¹ nad wodê"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "Modlitwa w Godzinie Mi³osierdzia Koronk¹ do Bo¿ego Mi³osierdzia"
        .Replacement.Text = "Koronka"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Przegl¹d katolickiego tygodnika ""Niedziela"""
        .Replacement.Text = "Przegl¹d tygodnika ""Niedziela"""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Warto zauwa¿yæ… w mijaj¹cym tygodniu"
        .Replacement.Text = "Warto zauwa¿yæ…"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Modlitwa z telefonicznym udzia³em dzieci"
        .Replacement.Text = "Modlitwa z udzia³em dzieci"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Modlitwa z telefonicznym udzia³em dzieci"
        .Replacement.Text = "Modlitwa z udzia³em dzieci"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Modlitwa z telefonicznym udzia³em dzieci"
        .Replacement.Text = "Modlitwa z udzia³em dzieci"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "Apel Jasnogórski z kaplicy Cudownego Obrazu Matki Bo¿ej Czêstochowskiej na Jasnej Górze"
        .Replacement.Text = "Apel Jasnogórski"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Katecheza ks. bp. Antoniego D³ugosza"
        .Replacement.Text = "Katecheza"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Jak my to widzimy - z daleka widaæ lepiej"
        .Replacement.Text = "Jak my to widzimy"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "0:00 Programy powtórkowe"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "Msza Œwiêta z kaplicy Cudownego Obrazu Matki Bo¿ej Czêstochowskiej na Jasnej Górze"
        .Replacement.Text = "Msza Œwiêta"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "Historia i architektura Polski w rysunkach prof. Ryszarda Natusiewicza"
        .Replacement.Text = "Historia i architektura Polski"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Audiencja Generalna Ojca Œwiêtego Franciszka z Watykanu"
        .Replacement.Text = "Audiencja Generalna"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub tvp_czyszczenie_dodatkowe()


 Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " (Yemin"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " (Elif"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Transmisja Mszy Œwiêtej z Sanktuarium Bo¿ego Mi³osierdzia w £agiewnikach - (JM), Transmisja"
        .Replacement.Text = "Transmisja Mszy Œwiêtej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "s.VI "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
       Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "s.VI "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "997) -"
        .Replacement.Text = "997 -"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
   
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^#/^#^#^#) - teleturniej"
        .Replacement.Text = "teleturniej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^#^#^#^#) - teleturniej"
        .Replacement.Text = "- teleturniej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Ko³o fortuny (^#^#^# ed. ^#) - teleturniej"
        .Replacement.Text = "Ko³o fortuny - teleturniej"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub
Sub hp_konerwsja_do_gazety()
'
' hp_konerwsja_do_gazety Makro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 10
        .Bold = True
        .Name = "Georgia"
    End With
    With Selection.Find.Replacement.ParagraphFormat
 
        .Alignment = wdAlignParagraphCenter
    End With
    
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue

    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = wdAlignParagraphJustify
    End With
    With Selection.Find.ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorBlack
            .BackgroundPatternColor = wdColorBlack
        End With
        '.Borders.Shadow = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 8
        .Italic = True
        .Name = "Georgia"
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = wdAlignParagraphJustify
    End With
    With Selection.Find.Replacement.ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorBlack
            .BackgroundPatternColor = wdColorBlack
        End With
       ' .Borders.Shadow = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = False
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
    With Selection.Find.ParagraphFormat
        With .Shading
            '.Texture = wdTextureNone
            .ForegroundPatternColor = wdColorBlack
            .BackgroundPatternColor = wdColorBlack
        End With
       ' .Borders.Shadow = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 8
        .Italic = True
        .Name = "Times New Roman"
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = wdAlignParagraphJustify
    End With
    With Selection.Find.Replacement.ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorBlack
            .BackgroundPatternColor = wdColorBlack
        End With
        '.Borders.Shadow = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'Interlinia 0,9
    
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(0.9)
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    
    
    'Sub FOTO()
             
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
                .Font.AllCaps = True
                With .ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(0.9)
                    .Alignment = wdAlignParagraphJustify
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                End With
                .Collapse 0
            End With
             
            'Exit Do
        Loop
    End With
lblf_Exit:
    Set fRng = Nothing
    'Exit Sub
    

End Sub

    
Sub hp_do_publico_xml()
'
' hp_do_publico_xml Makro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 10
        .Bold = True
        .Name = "Georgia"
    End With
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = wdAlignParagraphCenter
    End With
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 8
        .Italic = True
        .Name = "Times New Roman"
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 8
        .Bold = True
        .Name = "Times New Roman"
    End With
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = wdAlignParagraphCenter
    End With
    
    Selection.Find.Replacement.ClearFormatting
   
   
    With Selection.Find
        .Font.Size = 8
        .Font.Italic = True
        .Font.Name = "Times New Roman"
        .Font.Bold = True
        .Text = "^p"
        .Replacement.Text = ". "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


