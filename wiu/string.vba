Option Compare Database   'Use database order for string comparisons
Option Explicit

Function LowerCC()
' Converts the current control to lower case
  Screen.ActiveControl = LCase(Screen.ActiveControl)
End Function
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Parses address "New York NY 00123" into separate fields.
' Supports the following formats:
'   New York NY 12345-9876
'   Pierre, North Dakota 45678-7654
'   San Diego, CA, 98765-4321
'
' Words are extracted in the following order if no commas are found to delimit the values:
'   Zip, State, City
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Sub ParseCSZ(ByVal s As String, City As String, State As String, Zip As String)
Dim p As Double

' Check for comma after city name
  p = InStr(s, ",")
  If p > 0 Then
    City = Trim$(Left$(s, p - 1))
    s = Trim$(Mid$(s, p + 1))

'   Check for comma after state
    p = InStr(s, ",")
    If p > 0 Then
      State = Trim$(Left$(s, p - 1))
      Zip = Trim$(Mid$(s, p + 1))
    Else                           ' No comma between state and zip
      Zip = CutLastWord(s, s)
      State = s
    End If
  Else                             ' No commas between city, state, or zip
    Zip = CutLastWord(s, s)
    State = CutLastWord(s, s)
    City = s
  End If

' Clean up any dangling commas
  If Right$(State, 1) = "," Then
    State = Left$(State, Len(State) - 1)
  End If
  If Right$(City, 1) = "," Then
    City = Left$(City, Len(City) - 1)
  End If
End Sub

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Parses name "Mr. Bill A. Jones III, PhD" into separate fields.
' Words are extracted in the following order: Title, Degree, Pedigree, LName, FName, MName
' Assumes Pedigree is not preceded by a comma, or else it will end up with the Degree(s).
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Sub ParseName(ByVal s As String, Title As String, fName As String, MName As String, LName As String, Pedigree As String, Degree As String)
Dim Word As String, p As Integer, Found As Integer
Const Titles = "Mr.Mrs.Ms.Dr.Mme.Mssr.Mister,Miss,Doctor,Sir,Lord,Lady,Madam,Mayor,President"
Const Pedigrees = "Jr.Sr.III,IV,VIII,IX,XIII"
  Title = ""
  fName = ""
  MName = ""
  LName = ""
  Pedigree = ""
  Degree = ""
' Get Title
  'Word = CutWord(S, S)
  If InStr(Titles, Word) Then
    Title = Word
  Else
    s = Word & " " & s
  End If
  
  p = InStr(s, ",")
  If p > 0 Then
    Degree = Trim$(Mid$(s, p + 1))
    s = Trim$(Left$(s, p - 1))
  End If

' Get Pedigree
  Word = CutLastWord(s, s)
  If InStr(Pedigrees, Word) Then
    Pedigree = Word
  Else
    s = s & " " & Word
  End If
  LName = CutLastWord(s, s) ' Get Last Name
 'fName = CutWord(S, S)' Get First Name
  MName = Trim(s)
End Sub

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'  Capitalize first letter of every word in a field.
'  Use in an event procedure in AfterUpdate of control;
'  for example, [Last Name] = Proper([Last Name]).
'  Names such as O'Brien and Wilson-Smythe are properly capitalized,
'  but MacDonald is changed to Macdonald, and van Buren to Van Buren.
'  Note: For this function to work correctly, you must specify
'  Option Compare Database in the Declarations section of this module.
'
'  See Also: StrConv Function in the Microsoft Access 97 online Help.
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

Function Proper(x)
Dim temp$, C$, OldC$, I As Integer
  If IsNull(x) Then
    Exit Function
  Else
    temp$ = CStr(LCase(x))
    '  Initialize OldC$ to a single space because first
    '  letter needs to be capitalized but has no preceding letter.
    OldC$ = " "
    For I = 1 To Len(temp$)
      C$ = Mid$(temp$, I, 1)
      If C$ >= "a" And C$ <= "z" And (OldC$ < "a" Or OldC$ > "z") Then
        Mid$(temp$, I, 1) = UCase$(C$)
      End If
      OldC$ = C$
    Next I
    Proper = temp$
  End If
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Applies the Proper function to the current control
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function ProperCC()
  Screen.ActiveControl = Proper(Screen.ActiveControl)
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Similar to Proper(), but uses a table (NAMES) to look up words that don't fit the general formula.
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function ProperLookup(ByVal InText As Variant) As Variant
Dim OutText As String, Word As String, I As Integer, C As String
Dim db As Database, t As Recordset

' Output Null and other non-text as is
  If VarType(InText) <> 8 Then
    ProperLookup = InText
  Else
    Set db = CurrentDb
    Set t = db.OpenRecordset("Names", dbOpenTable)
    t.Index = "PrimaryKey"
    OutText = ""
    Word = ""
    For I = 1 To Len(InText)
      C = Mid$(InText, I, 1)
      Select Case C
        Case "A" To "Z"        ' if text, then build word
          Word = Word & C
        Case Else
          If Word <> "" Then   ' if not, then append existing word and then the character
            t.Seek "=", Word
            If t.NoMatch Then
              Word = UCase(Left(Word, 1)) & LCase(Mid(Word, 2))
            Else
              Word = t!name
            End If
            OutText = OutText & Word
            Word = ""
          End If
          OutText = OutText & C
      End Select
    Next I

' Process final word
    If Word <> "" Then
      t.Seek "=", Word
      If t.NoMatch Then
        Word = UCase(Left(Word, 1)) & LCase(Mid(Word, 2))
      Else
        Word = t!name
      End If
      OutText = OutText & Word
    End If

' Close table and return result
    t.Close
    db.Close
    ProperLookup = OutText
  End If
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Assumes N contains a single word
' N: can be null
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function ProperWord(N)
  ProperWord = UCase(Left(Trim(N), 1)) & LCase(Mid(Trim(N), 2))
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Converts a string to a series of hexadecimal digits.
' Useful if you want a true ASCII sort in your query.
' StrToHex(Chr(9) & "A~") returns "09417E"
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function StrToHex(s As Variant) As Variant
Dim temp As String, I As Integer
  If VarType(s) <> 8 Then
    StrToHex = s
  Else
    temp = ""
    For I = 1 To Len(s)
      temp = temp & Format(Hex(Asc(Mid(s, I, 1))), "00")
    Next I
    StrToHex = temp
  End If
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'  Use this procedure to test the ParseName and ParseCSZ procedures
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Sub TestParseName()
  Dim N As String, t As String, f As String, M As String, l As String, p As String, d As String
  N = "Dr. James George William Joyce-Brothers IV, MS, PhD"
  ParseName N, t, f, M, l, p, d
  'Debug.Print t, f, M, l, P, d
  N = "New York NY 45678-9876"
  ParseCSZ N, t, f, M
  'Debug.Print t, f, M
End Sub

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Àctive control
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function UpperCC()
  Screen.ActiveControl = UCase(Screen.ActiveControl)
End Function
