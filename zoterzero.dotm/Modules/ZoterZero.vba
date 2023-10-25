'' Original code ZoterZero4 by djross3 on
'' pastebin URL: https://pastebin.com/tSbsHg6w
'' Zotero forum URL: https://forums.zotero.org/discussion/comment/340209/#Comment_340209

Option Explicit

Private Sub getAuthorYear(ByVal str As String, ByRef var As Variant)
  ' This function is specific for APA style citation
  var = Split(str, ",") ' split at comma for (Author, YYYY, p. strArg), store as var
  var(0) = Replace(var(0), "(", "", 1, 1) ' first part is Author, remove "(" from Author
  var(1) = Split(var(1), " ")(1) ' split second part of var using " ", save second part of it as Year
End Sub

Private Function ZoterZeroFieldFix(F) ' Fix a field with ZoterZero
  Dim fieldChanged As Byte
  Dim fieldText As String
  fieldChanged = 0 ' assume nothing changed, unless...
  fieldText = F.Code.Text ' get the current code of the field as text
  If InStr(fieldText, " ADDIN ZOTERO_ITEM CSL_CITATION") = 0 Then Exit Function ' make sure this is a Zotero field to modify, starts with right text
  Dim myArray As Variant
  myArray = Split(fieldText, " ", 5) ' split that code into parts up to 5
  Dim Json As Dictionary ' create a Json object to work with
  Set Json = JsonConverter.ParseJson(myArray(4)) ' get the fourth part, which is the Json data
  If (Json("citationItems").Count > 1) Then Exit Function ' only works for one citation item
  Dim myResult As Variant
  myResult = F.Result.Text ' get the displayed text from the field to work with it below
  
  If Left(myResult, 1) <> "(" Then Exit Function ' only works if original citation starts with "("
  
  Dim strLocator As String
  strLocator = Json("citationItems")(1)("locator")
  
  Dim varResult As Variant
  getAuthorYear myResult, varResult ' split myResult into string array
  
  '' Locator switch
  Select Case strLocator
    Case "a" ' if the locator is set to "a" - author
      myResult = varResult(0)
      fieldChanged = 1 ' report updated below
    Case "a (y)" ' if the locator is set to "a (y)" - author (YYYY)
      myResult = varResult(0) & " (" & varResult(1) & ")"
      fieldChanged = 1 ' report updated below
    Case "a y" ' if the locator is set to "a y" - author YYYY
      myResult = varResult(0) & " " & varResult(1)
      fieldChanged = 1 ' report updated below
    Case "y" ' if the locator is set to "y" - YYYY *CAUTION: do not check Omit author*
      myResult = varResult(1)
      fieldChanged = 1 ' report updated below
  End Select

  '' SECONDARY TEXT REPLACEMENT FOR DOUBLED PARENTHESES TO REMOVE PARENTHESES
  If Left(myResult, 2) = "((" Then ' if it begins with doubled open parentheses...
      myResult = Replace(myResult, "((", "", , 1) ' replace "((" with "" once
      fieldChanged = 1 ' report updated below
  End If
  If Right(myResult, 2) = "))" Then ' if it ends with doubled open parentheses...
      myResult = Replace(myResult, "))", "", , 1) ' replace "))" with "" once
      fieldChanged = 1 ' report updated below
  End If

  '' CAPITALIZE FIRST LETTER IF NEEDED FOR UNUSUAL CASES LIKE "von" > "Von" SENTENCE INITIALLY
  If Left(myResult, 1) = "^" Then ' if it begins with a carrot marker...
    Dim myResultChar As String
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
    Dim myJson As String
    
    ' the following two lines make Zotero think this was the original output, so no warnings!
    Json("properties")("plainCitation") = myResult ' set the Json citation data to new label
    Json("properties")("formattedCitation") = myResult ' again, other instance
    
    F.Result.Text = myResult ' replace the displayed text with the new text
    F.Result.Font.Underline = wdUnderlineNone ' remove dashed underlining from Zotero's delayed update feature if present
    myJson = JsonConverter.ConvertToJson(Json) ' collapse Json back to text
    F.Code.Text = " " & myArray(1) & " " & myArray(2) & " " & myArray(3) & " " & myJson & " " ' reconstruct field code
    ZoterZeroFieldFix = 1 ' updated, return success
  End If
End Function

Sub ZoterZero()
  Dim changeSuccess As Byte
  Dim checkField As Integer
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
                checkField = oShp.TextFrame.TextRange.Fields.Count ' get the total number of fields in this section
                While checkField > 0 ' check each field
                  changeSuccess = ZoterZeroFieldFix(oShp.TextFrame.TextRange.Fields(checkField)) ' check and fix this field
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