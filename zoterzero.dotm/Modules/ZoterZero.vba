'' Original code ZoterZero4 by djross3 on
'' pastebin URL: https://pastebin.com/tSbsHg6w
'' Zotero forum URL: https://forums.zotero.org/discussion/comment/340209/#Comment_340209

Function ZoterZeroFieldFix(F) ' Fix a field with ZoterZero
    fieldChanged = 0 ' assume nothing changed, unless...
    fieldText = F.Code.Text ' get the current code of the field as text
    If InStr(fieldText, " ADDIN ZOTERO_ITEM CSL_CITATION") = 1 Then ' make sure this is a Zotero field to modify, starts with right text
        myArray = Split(fieldText, " ", 5) ' split that code into parts up to 5
        Dim Json As Object ' create a Json object to work with
        Set Json = JsonConverter.ParseJson(myArray(4)) ' get the fourth part, which is the Json data
        myResult = F.Result.Text ' get the displayed text from the field to work with it below
        '' MAIN TEXT REPLACEMENT OF PAGE "0" WITH AUTHOR NAME ONLY
        If Json("citationItems").Count = 1 Then ' only works for one citation item
            If Json("citationItems")(1)("locator") = "0" Then ' if the page range is set to "0" process for ZoterZero:
               If InStr(myResult, "(") = 1 Then ' if it begins with open parentheses...
                    myResult = Split(myResult, "(", 2) ' split at parentheses, only 2 parts
                    myResult = myResult(1) ' get the second part (without open parentheses)
                    myResult = StrReverse(myResult) ' reverse the string so we can work from the other end:
                    splitChar = " " ' default character to split is a space, occurring between name and year
                    If InStr(myArray(4), """issued"":") = 0 Then ' if the year is missing, allow year-less name-only cites without parentheses:
                        splitChar = ":" ' instead split at the colon because there is no year
                    End If
                    myResult = Split(myResult, splitChar, 2) ' split at splitChar, up to 2 parts
                    myResult = myResult(1) ' get the second part (without the actually-last [=year] section)
                    myResult = StrReverse(myResult) ' reverse back to normal order
                    fieldChanged = 1 ' report updated below
                End If
            '' ALTERNATIVE TEXT REPLACEMENT OF PAGE "00" WITH "AUTHOR (YYYY)" FORMAT, **ONLY FOR MOST COMMON SIMPLE SCENARIO**
            ''' LIMITATION: ONLY WORKS WITHOUT PAGE NUMBERS DUE TO OBVIOUS CONFLICT WITH "00"...
            ElseIf Json("citationItems")(1)("locator") = "00" Then ' if the page range is set to "0" process for ZoterZero:
               If InStr(myResult, "(") = 1 And InStr(myArray(4), """issued"":") Then ' if it begins with open parentheses and contains a date...
                    myResult = Split(myResult, "(", 2) ' split at parentheses, only 2 parts
                    myResult = myResult(1) ' get the second part (without open parentheses)
                    myResult = StrReverse(myResult) ' reverse the string so we can work from the other end:
                    splitChar = " " ' default character to split is a space, occurring between name and year
                    myResult = Split(myResult, splitChar, 2) ' split at splitChar, up to 2 parts
                    myResult = myResult(0) & "(" & splitChar & myResult(1) ' recombine with parentheses inserted
                    myResult = Split(myResult, ":", 2) ' split without pages, up to 2 parts
                    myResult = ")" & myResult(1) ' recombine without pages
                    myResult = StrReverse(myResult) ' reverse back to normal order
                    fieldChanged = 1 ' report updated below
                End If
            End If
        End If
        '' SECONDARY TEXT REPLACEMENT FOR DOUBLED PARENTHESES TO REMOVE PARENTHESES
        If Left(myResult, 2) = "((" Then ' if it begins with doubled open parentheses...
            myResult = Split(myResult, "((", 2) ' split at double parentheses, only 2 parts
            myResult = myResult(1) ' get the second part (without open double parentheses)
            fieldChanged = 1 ' report updated below
        End If
        If Right(myResult, 2) = "))" Then ' if it ends with doubled open parentheses...
            myResult = StrReverse(myResult) ' reverse the string so we can work from the end:
            myResult = Split(myResult, "))", 2) ' split at double parentheses, only 2 parts
            myResult = myResult(1) ' get the second part (without close double parentheses)
            myResult = StrReverse(myResult) ' reverse the string back to normal:
            fieldChanged = 1 ' report updated below
        End If
        '' CAPITALIZE FIRST LETTER IF NEEDED FOR UNUSUAL CASES LIKE "von" > "Von" SENTENCE INITIALLY
        If Left(myResult, 1) = "^" Then ' if it begins with a carrot marker...
            myResult = Split(myResult, "^", 2) ' split at carrot, only 2 parts
            myResult = myResult(1) ' get the second part (without carrot)
            myResultChar = Left(myResult, 1) ' save first character
            myResult = Split(myResult, Left(myResult, 1), 2) ' split at first character, only 2 parts
            myResult = myResult(1) ' save rest, missing first character
            myResult = UCase(myResultChar) & myResult ' combine uppercase version of first character and rest
            fieldChanged = 1 ' report updated below
        End If
        '' UPDATE AND SAVE FIELD IF CHANGED ABOVE:
        If fieldChanged = 1 Then ' the field text has been updated as MyResult, let's save it:
            ' the following two lines make Zotero think this was the original output, so no warnings!
            Json("properties")("plainCitation") = myResult ' set the Json citation data to new label
            Json("properties")("formattedCitation") = myResult ' again, other instance
            F.Result.Text = myResult ' replace the displayed text with the new text
            F.Result.Font.Underline = wdUnderlineNone ' remove dashed underlining from Zotero's delayed update feature if present
            myJson = JsonConverter.ConvertToJson(Json) ' collapse Json back to text
            F.Code.Text = " " & myArray(1) & " " & myArray(2) & " " & myArray(3) & " " & myJson & " " ' reconstruct field code
            ZoterZeroFieldFix = 1 ' updated, return success
        End If
    End If
End Function
 
Sub ZoterZero()
'
' ZoterZero main function
'' if selection or text near cursor contains fields, check and fix them
'' else check and fix all fields in document
changeSuccess = 0 ' no fields fixed yet
Selection.Expand Unit:=wdSentence ' expand the selection to at least sentence-level
If Selection.Fields.Count > 0 Then ' if fields are selected...
    checkField = Selection.Fields.Count ' get the total number of fields
    While checkField > 0 ' check each field
        changeSuccess = ZoterZeroFieldFix(Selection.Fields(checkField)) ' check and fix this field
        checkField = checkField - 1 ' check the previous field next
    Wend
End If
If changeSuccess = 0 Then ' no fields have been updated yet, let's update all fields in document
        ' based on http://www.vbaexpress.com/kb/getarticle.php?kb_id=1100
        Dim rngStory As Word.Range ' vars for below
        Dim lngValidate As Long ' vars for below
        Dim oShp As Shape ' vars for below
        lngValidate = ActiveDocument.Sections(1).Headers(1).Range.StoryType ' starting point
        For Each rngStory In ActiveDocument.StoryRanges 'Iterate through all linked stories
          Do
            On Error Resume Next
            checkField = rngStory.Fields.Count ' get the total number of fields in this section
            While checkField > 0 ' check each field
                changeSuccess = ZoterZeroFieldFix(rngStory.Fields(checkField)) ' check and fix this field
                checkField = checkField - 1 ' check the previous field next
            Wend
            Select Case rngStory.StoryType
              Case 6, 7, 8, 9, 10, 11
                If rngStory.ShapeRange.Count > 0 Then
                  For Each oShp In rngStory.ShapeRange
                    If oShp.TextFrame.HasText Then
                        checkField = Shp.TextFrame.TextRange.Fields.Count ' get the total number of fields in this section
                        While checkField > 0 ' check each field
                            changeSuccess = ZoterZeroFieldFix(Shp.TextFrame.TextRange.Fields(checkField)) ' check and fix this field
                            checkField = checkField - 1 ' check the previous field next
                        Wend
                    End If
                  Next
                End If
              Case Else 'Do Nothing
            End Select
            On Error GoTo 0
            'Get next linked story (if any)
            Set rngStory = rngStory.NextStoryRange ' get ready for next section
          Loop Until rngStory Is Nothing ' keep going through until all sections are done
        Next
    End If
    Selection.Collapse ' reset cursor to beginning of section which isn't quite right but close enough
End Sub