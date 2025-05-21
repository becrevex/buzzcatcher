Sub Buzzcatcher()
    Dim phrases As Variant
    Dim i As Long
    Dim rng As Range
    Dim phrase As String

    phrases = Array( _
        "At the end of the day", "With that being said", "It goes without saying", "In a nutshell", _
        "Needless to say", "When it comes to", "A significant number of", "It’s worth mentioning", _
        "Last but not least", "Cutting‑edge", "Leveraging", "Moving forward", "Going forward", _
        "On the other hand", "Notwithstanding", "Takeaway", "As a matter of fact", "In the realm of", _
        "Seamless integration", "Robust framework", "Holistic approach", "Paradigm shift", "Synergy", _
        "Scale-up", "Optimize", "Game‑changer", "Unleash", "Uncover", "In a world", "In a sea of", _
        "Digital landscape", "Elevate", "Embark", "Delve", "Game Changer", "In the midst", "In addition" _
    )
    
    For i = LBound(phrases) To UBound(phrases)
        phrase = phrases(i)
        
        Set rng = ActiveDocument.Content
        With rng.Find
            .ClearFormatting
            .Text = phrase
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute
                rng.HighlightColorIndex = wdYellow
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next i

    MsgBox "Execution complete!", vbInformation
End Sub
