Sub XpressPunte4_0UNO()
Attribute XpressPunte4_0UNO.VB_Description = "plantilla"
Attribute XpressPunte4_0UNO.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' XpressPunte4_0 Macro
' Cuarta versión de esta herramienta.
'
    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.Italic = True
    
    With Selection.Find
        .Text = "examen"
        .Replacement.Text = "examen"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "más"
        .Replacement.Text = "más"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = Tr
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "menos"
        .Replacement.Text = "menos"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = Tru
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "importante"
        .Replacement.Text = "importante"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "importancia"
        .Replacement.Text = "importancia"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = Tru
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "esencial"
        .Replacement.Text = "esencial"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "esencialmente"
        .Replacement.Text = "esencialmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "frecuente"
        .Replacement.Text = "frecuente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "frecuentemente"
        .Replacement.Text = "frecuentemente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "infrecuente"
        .Replacement.Text = "infrecuente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "infrecuentemente"
        .Replacement.Text = "infrecuentemente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "excepcional"
        .Replacement.Text = "excepcional"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ojo"
        .Replacement.Text = "ojo"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "!"
        .Replacement.Text = "!"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "¡"
        .Replacement.Text = "¡"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "no"
        .Replacement.Text = "no"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "sí"
        .Replacement.Text = "sí"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "no"
        .Replacement.Text = "no"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "normalmente"
        .Replacement.Text = "normalmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "habitual"
        .Replacement.Text = "habitual"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "habitualmente"
        .Replacement.Text = "habitualmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "común"
        .Replacement.Text = "común"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "comúnmente"
        .Replacement.Text = "comúnmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "comunmente"
        .Replacement.Text = "comunmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "principal"
        .Replacement.Text = "principal"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "principalmente"
        .Replacement.Text = "principalmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "típica"
        .Replacement.Text = "típica"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "típicamente"
        .Replacement.Text = "típicamente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "suele"
        .Replacement.Text = "suele"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "destaca"
        .Replacement.Text = "destaca"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "destacar"
        .Replacement.Text = "destacar"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "destacable"
        .Replacement.Text = "destacable"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "sobresale"
        .Replacement.Text = "sobresale"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "sobresalir"
        .Replacement.Text = "sobresalir"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "remarcable"
        .Replacement.Text = "remarcable"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "mayoría"
        .Replacement.Text = "mayoría"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "mayormente"
        .Replacement.Text = "mayormente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "mayoritaria"
        .Replacement.Text = "mayoritaria"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "mayoritariamente"
        .Replacement.Text = "mayoritariamente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "50%"
        .Replacement.Text = "50%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "51%"
        .Replacement.Text = "51%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "52%"
        .Replacement.Text = "52%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "53%"
        .Replacement.Text = "53%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "54%"
        .Replacement.Text = "54%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "55%"
        .Replacement.Text = "55%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "56%"
        .Replacement.Text = "56%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "57%"
        .Replacement.Text = "57%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "58%"
        .Replacement.Text = "58%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "59%"
        .Replacement.Text = "59%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "60%"
        .Replacement.Text = "60%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "61%"
        .Replacement.Text = "61%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "62%"
        .Replacement.Text = "62%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "63%"
        .Replacement.Text = "63%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "64%"
        .Replacement.Text = "64%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "65%"
        .Replacement.Text = "65%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "66%"
        .Replacement.Text = "66%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "67%"
        .Replacement.Text = "67%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "68%"
        .Replacement.Text = "68%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "69%"
        .Replacement.Text = "69%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "70%"
        .Replacement.Text = "70%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "71%"
        .Replacement.Text = "71%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "72%"
        .Replacement.Text = "72%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "73%"
        .Replacement.Text = "73%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "74%"
        .Replacement.Text = "74%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "75%"
        .Replacement.Text = "75%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "76%"
        .Replacement.Text = "76%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "77%"
        .Replacement.Text = "77%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "78%"
        .Replacement.Text = "78%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "79%"
        .Replacement.Text = "79%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "80%"
        .Replacement.Text = "80%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "81%"
        .Replacement.Text = "81%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "82%"
        .Replacement.Text = "82%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "83%"
        .Replacement.Text = "83%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "84%"
        .Replacement.Text = "84%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "85%"
        .Replacement.Text = "85%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "86%"
        .Replacement.Text = "86%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "87%"
        .Replacement.Text = "87%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "88%"
        .Replacement.Text = "88%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "89%"
        .Replacement.Text = "89%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "90%"
        .Replacement.Text = "90%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "91%"
        .Replacement.Text = "91%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "92%"
        .Replacement.Text = "92%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "93%"
        .Replacement.Text = "93%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "94%"
        .Replacement.Text = "94%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "95%"
        .Replacement.Text = "95%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "96%"
        .Replacement.Text = "96%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "97%"
        .Replacement.Text = "97%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "98%"
        .Replacement.Text = "98%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "99%"
        .Replacement.Text = "99%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "100%"
        .Replacement.Text = "100%"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "particular"
        .Replacement.Text = "particular"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "particularmente"
        .Replacement.Text = "particularmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "propio"
        .Replacement.Text = "propio"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "propia"
        .Replacement.Text = "propia"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "característico"
        .Replacement.Text = "característico"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "característica"
        .Replacement.Text = "característica"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "distintivo"
        .Replacement.Text = "distintivo"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "distintiva"
        .Replacement.Text = "distintiva"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "estudiar"
        .Replacement.Text = "estudiar"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "recuerda"
        .Replacement.Text = "recuerda"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "recordar"
        .Replacement.Text = "recordar"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "recordad"
        .Replacement.Text = "recordad"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "fijo"
        .Replacement.Text = "fijo"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "sobre todo"
        .Replacement.Text = "sobre todo"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "todo"
        .Replacement.Text = "todo"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "muy"
        .Replacement.Text = "muy"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "toda"
        .Replacement.Text = "toda"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "todos"
        .Replacement.Text = "todos"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "todas"
        .Replacement.Text = "todas"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "fundamental"
        .Replacement.Text = "fundamental"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "fundamentalmente"
        .Replacement.Text = "fundamentalmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "obligatorio"
        .Replacement.Text = "obligatorio"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "obligatoriamente"
        .Replacement.Text = "obligatoriamente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "obligatoria"
        .Replacement.Text = "obligatoria"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "generalmente"
        .Replacement.Text = "generalmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "a menudo"
        .Replacement.Text = "a menudo"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "MIR"
        .Replacement.Text = "MIR"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "especialmente"
        .Replacement.Text = "especialmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "rara"
        .Replacement.Text = "rara"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "raramente"
        .Replacement.Text = "raramente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "caracterizado"
        .Replacement.Text = "caracterizado"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "caracterizada"
        .Replacement.Text = "caracterizada"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "caracterizados"
        .Replacement.Text = "caracterizados"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "caracterizadas"
        .Replacement.Text = "caracterizadas"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "grave"
        .Replacement.Text = "grave"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "morir"
        .Replacement.Text = "morir"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "muerte"
        .Replacement.Text = "muerte"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "fallece"
        .Replacement.Text = "fallece"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "fallecimiento"
        .Replacement.Text = "fallecimiento"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "raro"
        .Replacement.Text = "raro"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "de elección"
        .Replacement.Text = "de elección"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub XpressPunte4_0DOS()

    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.Italic = True
    
    With Selection.Find
        .Text = "examen"
        .Replacement.Text = "examen"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "50 %"
        .Replacement.Text = "50 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "51 %"
        .Replacement.Text = "51 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "52 %"
        .Replacement.Text = "52 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "53 %"
        .Replacement.Text = "53 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "54 %"
        .Replacement.Text = "54 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "55 %"
        .Replacement.Text = "55 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "56 %"
        .Replacement.Text = "56 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "57 %"
        .Replacement.Text = "57 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "58 %"
        .Replacement.Text = "58 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "59 %"
        .Replacement.Text = "59 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "60 %"
        .Replacement.Text = "60 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "61 %"
        .Replacement.Text = "61 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "62 %"
        .Replacement.Text = "62 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "63 %"
        .Replacement.Text = "63 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "64 %"
        .Replacement.Text = "64 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "65 %"
        .Replacement.Text = "65 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "66 %"
        .Replacement.Text = "66 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "67 %"
        .Replacement.Text = "67 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "68 %"
        .Replacement.Text = "68 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "69 %"
        .Replacement.Text = "69 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "70 %"
        .Replacement.Text = "70 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "71 %"
        .Replacement.Text = "71 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "72 %"
        .Replacement.Text = "72 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "73 %"
        .Replacement.Text = "73 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "74 %"
        .Replacement.Text = "74 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "75 %"
        .Replacement.Text = "75 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "76 %"
        .Replacement.Text = "76 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "77 %"
        .Replacement.Text = "77 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "78 %"
        .Replacement.Text = "78 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "79 %"
        .Replacement.Text = "79 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "80 %"
        .Replacement.Text = "80 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "81 %"
        .Replacement.Text = "81 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "82 %"
        .Replacement.Text = "82 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "83 %"
        .Replacement.Text = "83 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "84 %"
        .Replacement.Text = "84 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "85 %"
        .Replacement.Text = "85 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "86 %"
        .Replacement.Text = "86 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "87 %"
        .Replacement.Text = "87 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "88 %"
        .Replacement.Text = "88 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "89 %"
        .Replacement.Text = "89 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "90 %"
        .Replacement.Text = "90 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "91 %"
        .Replacement.Text = "91 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "92 %"
        .Replacement.Text = "92 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "93 %"
        .Replacement.Text = "93 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "94 %"
        .Replacement.Text = "94 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "95 %"
        .Replacement.Text = "95 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "96 %"
        .Replacement.Text = "96 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "97 %"
        .Replacement.Text = "97 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "98 %"
        .Replacement.Text = "98 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "99 %"
        .Replacement.Text = "99 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "100 %"
        .Replacement.Text = "100 %"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "abundante"
        .Replacement.Text = "abundante"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "alta"
        .Replacement.Text = "alta"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "altísima"
        .Replacement.Text = "altísima"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "asintomático"
        .Replacement.Text = "asintomático"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "aunque"
        .Replacement.Text = "aunque"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "baja"
        .Replacement.Text = "baja"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "bajísima"
        .Replacement.Text = "bajísima"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "característicamente"
        .Replacement.Text = "característicamente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "característico"
        .Replacement.Text = "característico"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "común"
        .Replacement.Text = "común"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "debemos"
        .Replacement.Text = "debemos"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "define"
        .Replacement.Text = "define"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "desconocida"
        .Replacement.Text = "desconocida"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "disminuida"
        .Replacement.Text = "disminuida"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "dudosa"
        .Replacement.Text = "dudosa"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "elevada"
        .Replacement.Text = "elevada"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "es característica"
        .Replacement.Text = "es característica"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "es característico"
        .Replacement.Text = "es característico"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "escaso"
        .Replacement.Text = "escaso"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "específica"
        .Replacement.Text = "específica"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "específicas"
        .Replacement.Text = "específicas"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "específico"
        .Replacement.Text = "específico"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "específicos"
        .Replacement.Text = "específicos"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "excepción"
        .Replacement.Text = "excepción"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "excepcional"
        .Replacement.Text = "excepcional"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "excepcionales"
        .Replacement.Text = "excepcionales"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "excepciones"
        .Replacement.Text = "excepciones"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "excepto"
        .Replacement.Text = "excepto"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "exclusivamente"
        .Replacement.Text = "exclusivamente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "exclusivo"
        .Replacement.Text = "exclusivo"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "exige"
        .Replacement.Text = "exige"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "frecuente"
        .Replacement.Text = "frecuente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "gold standart"
        .Replacement.Text = "gold standart"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "gold standard"
        .Replacement.Text = "gold standard"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "habitual"
        .Replacement.Text = "habitual"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "habitualmente"
        .Replacement.Text = "habitualmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "habría que"
        .Replacement.Text = "habría que"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ideal"
        .Replacement.Text = "ideal"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "idealmente"
        .Replacement.Text = "idealmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "IMP"
        .Replacement.Text = "IMP"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "importantes"
        .Replacement.Text = "importantes"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "indeterminada"
        .Replacement.Text = "indeterminada"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "indispensable"
        .Replacement.Text = "indispensable"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "lo primero"
        .Replacement.Text = "lo primero"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "en primer lugar"
        .Replacement.Text = "en primer lugar"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "mal"
        .Replacement.Text = "mal"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "mayor"
        .Replacement.Text = "mayor"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "mucho"
        .Replacement.Text = "mucho"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "necesaridad"
        .Replacement.Text = "necesaridad"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "necesario"
        .Replacement.Text = "necesario"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "necesita"
        .Replacement.Text = "necesita"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "normalmente"
        .Replacement.Text = "normalmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "obligado"
        .Replacement.Text = "obligado"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "obligatoriamente"
        .Replacement.Text = "obligatoriamente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "obligatoriedad"
        .Replacement.Text = "obligatoriedad"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "obligatorio"
        .Replacement.Text = "obligatorio"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "patognomónico"
        .Replacement.Text = "patognomónico"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "patrón oro"
        .Replacement.Text = "patrón oro"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "pero"
        .Replacement.Text = "pero"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "poco"
        .Replacement.Text = "poco"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "precisa"
        .Replacement.Text = "precisa"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "prevalece"
        .Replacement.Text = "prevalece"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "prima"
        .Replacement.Text = "prima"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "primero"
        .Replacement.Text = "primero"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "principal"
        .Replacement.Text = "principal"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "principalmente"
        .Replacement.Text = "principalmente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "problema"
        .Replacement.Text = "problema"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "se beneficia"
        .Replacement.Text = "se beneficia"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "se debería"
        .Replacement.Text = "se debería"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "se habría que"
        .Replacement.Text = "se habría que"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "se tendría que"
        .Replacement.Text = "se tendría que"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "siempre"
        .Replacement.Text = "siempre"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "síndrome"
        .Replacement.Text = "síndrome"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "sólamente"
        .Replacement.Text = "sólamente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "sólo"
        .Replacement.Text = "sólo"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "tan"
        .Replacement.Text = "tan"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "tendríamos que"
        .Replacement.Text = "tendríamos que"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "tenemos que"
        .Replacement.Text = "tenemos que"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "únicamente"
        .Replacement.Text = "únicamente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "único"
        .Replacement.Text = "único"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "única"
        .Replacement.Text = "única"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "urgente"
        .Replacement.Text = "urgente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "emergente"
        .Replacement.Text = "emergente"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "emergencia"
        .Replacement.Text = "emergencia"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "urgencia"
        .Replacement.Text = "urgencia"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
