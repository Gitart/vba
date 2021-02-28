Option Compare Database   'Use database order for string comparisons
Option Explicit

'C÷èòàåò êîëè÷åñòâî ñëîâ ðàçäåëåííûx çàïÿòîé
Function CountCSVWords(s) As Double
Dim WC As Integer, Pos As Integer
  If VarType(s) <> 8 Or Len(s) = 0 Then
    CountCSVWords = 0
    Exit Function
  End If
  WC = 1
  Pos = InStr(s, ",")
  Do While Pos > 0
    WC = WC + 1
    Pos = InStr(Pos + 1, s, ",")
  Loop
  CountCSVWords = WC
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'êîëè÷ ñëîâ ðàçäåëåííûå ïðîáåëàìè
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function CountWords(s) As Double
Dim WC As Double, I As Double, OnASpace As Double
  If VarType(s) <> 8 Or Len(Trim(s)) = 0 Then
    CountWords = 0
    Exit Function
  End If
  WC = 0
  OnASpace = True
  For I = 1 To Len(s)
    If Mid(s, I, 1) = " " Then
      OnASpace = True
    Else
      If OnASpace Then
        OnASpace = False
        WC = WC + 1
      End If
    End If
  Next I
  CountWords = WC
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'âîçâðàùàåò ïåðâîå ñëîâî â ïðåäëîæåíèè
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function CutFirstWord(s, Remainder)
Dim temp, I As Double, p As Double
  temp = Trim(s)
  p = InStr(temp, " ")
  If p = 0 Then
    CutFirstWord = temp
    Remainder = Null
  Else
    CutFirstWord = Left(temp, p - 1)
    Remainder = Trim(Mid(temp, p + 1))
  End If
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'âîçâðàùàåò ïîñëåäíåå ñëîâî â ïðåäëîæåíèè
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function CutLastWord(s, Remainder)
Dim temp, I As Double, p As Double
  temp = Trim(s)
  p = 1
  For I = Len(temp) To 1 Step -1
    If Mid(temp, I, 1) = " " Then
      p = I + 1
      Exit For
    End If
  Next I
  If p = 1 Then
    CutLastWord = temp
    Remainder = Null
  Else
    CutLastWord = Mid(temp, p)
    Remainder = Trim(Left(temp, p - 1))
  End If
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'âîçâðàùàåò ñëîâî â ïðåäëîæåíèè ïî íîìåðó èíäåêñà âî âòîðì àðãóìåíòå
'ñëîâà ðàçäåëåíû çàïÿòîé
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function GetCSVWord(s, Indx As Double)
Dim WC As Double, Count As Double, SPos As Double, EPos As Double
  WC = CountCSVWords(s)
  If Indx < 1 Or Indx > WC Then
    GetCSVWord = Null
    Exit Function
  End If
  Count = 1
  SPos = 1
  For Count = 2 To Indx
    SPos = InStr(SPos, s, ",") + 1
  Next Count
  EPos = InStr(SPos, s, ",") - 1
  If EPos <= 0 Then EPos = Len(s)
  GetCSVWord = Mid(s, SPos, EPos - SPos + 1)
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'âîçâðàùàåò ñëîâî èç ïðåäëîæåíèÿ ðàçäåëåííîããî ïðîáåëàìè
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function GetWord(s, Indx As Double)
Dim I As Double, WC As Double, Count As Double, SPos As Double, EPos As Double, OnASpace As Double
  WC = CountWords(s)
  If Indx < 1 Or Indx > WC Then
    GetWord = Null
    Exit Function
  End If
  Count = 0
  OnASpace = True
  For I = 1 To Len(s)
    If Mid(s, I, 1) = " " Then
      OnASpace = True
    Else
      If OnASpace Then
        OnASpace = False
        Count = Count + 1
        If Count = Indx Then
          SPos = I
          Exit For
        End If
      End If
    End If
  Next I
  EPos = InStr(SPos, s, " ") - 1
  If EPos <= 0 Then EPos = Len(s)
  GetWord = Mid(s, SPos, EPos - SPos + 1)
End Function
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Åñëè âûðàæåíèå ñîâïàäàåò ñ ìàñêîé òî òðóå èíà÷å ôàëüø
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function Like2(ByVal Text As String, ByVal Mask As String) As Double
'
' This function does simple pattern matching.
' It allows the following wildcards:
'   # (digit)
'   ? (any character)
'   @ (alpha)
'
Dim Match As Double, I As Double, C As String * 1, MC As String * 1
  If Len(Text) <> Len(Mask) Then
    Match = False
  Else
    Match = True
    For I = 1 To Len(Mask)
      C = Mid(Text, I, 1)
      MC = Mid(Mask, I, 1)
      Select Case MC
        Case "#"  ' Match digit
          If C < "0" Or C > "9" Then
            Match = False
            Exit For
          End If
        Case "@"  ' Match A-Z
          If Not (C >= "A" And C <= "Z") And Not (C >= "a" And C <= "z") Then
            Match = False
            Exit For
          End If
        Case "?"  ' Match anything
        Case Else ' Exact match
          If C <> MC Then
            Match = False
            Exit For
          End If
      End Select
    Next I
  End If
  Like2 = Match
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'lpad("tt",".",15)  -> .............tt
' Adds character C to the left of S to make it right-justified
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function LPad(s, ByVal C As String, N As Double) As String
  If Len(C) = 0 Then C = " "
  If N < 1 Then
    LPad = ""
  Else
    LPad = Right$(String$(N, Left$(C, 1)) & s, N)
  End If
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Function parses S, using delimiter Delim, and copies each element into array A().
' The function returns the number of items copied.
' Compare:
'   0 = Binary comparison     - can search for Tabs (chr$(9))
'   1 = Text comparison       - can't search for Tabs
'   2 = Database comparison   - can't search for Tabs
'
' Calling convention:
'   ReDim Items(20)
'   ItemCount = ParseItemsToArray("A,B,C",Items(),Delim)
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

Function ParseItemsToArray(ByVal s As String, a() As String, ByVal Delim As String, ByVal Compare As Double) As Double
Dim p As Double, I As Double
  If Delim = "" Then
    ParseItemsToArray = -1
    Exit Function
  End If
'
' Copy Items
'
  I = 0
  p = InStr(1, s, Delim, Compare)
  Do While p > 0
    a(LBound(a) + I) = Left$(s, p - 1)
    I = I + 1
    s = Mid$(s, p + 1)
    p = InStr(1, s, Delim, Compare)
  Loop
'
' Copy Last Item
'
  a(LBound(a) + I) = s
  I = I + 1
'
  ParseItemsToArray = I
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Replaces the SearchStr string with Replacement string in the TextIn string.
' Uses CompMode to determine comparison mode
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

Function ReplaceStr(TextIn, SearchStr, Replacement, CompMode As Double)
Dim WorkText As String, Pointer As Double
  If IsNull(TextIn) Then
    ReplaceStr = Null
  Else
    WorkText = TextIn
    Pointer = InStr(1, WorkText, SearchStr, CompMode)
    Do While Pointer > 0
      WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
      Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr, CompMode)
    Loop
    ReplaceStr = WorkText
  End If
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Adds character C to the right of S to make it left-justified.
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function RPad(s, ByVal C As String, N As Double) As String
If Len(C) = 0 Then C = " "
  If N < 1 Then
    RPad = ""
  Else
    RPad = Left$(s & String$(N, Left$(C, 1)), N)
  End If
End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Use this procedure to test the ParseItemsToArray procedure
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Sub TestParseItems()
Dim ItemCount As Double, I As Double
ReDim ItemArray(1 To 20) As String
  ItemCount = ParseItemsToArray("A,B,C", ItemArray(), ",", 0)
  For I = LBound(ItemArray) To LBound(ItemArray) + ItemCount - 1
    Debug.Print ItemArray(I)
  Next I
End Sub

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
' Removes articles (a, an, the) from the beginning of a string.
' If you specify TRUE for the varKeepArticle argument, the article is
' moved to the end of the string.
' ParseArticle("The Beatles") returns "Beatles."
' ParseArticle("The Beatles", True) returns "Beatles, The."
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Function ParseArticle(strOldTitle As String, Optional varKeepArticle As Variant) As String
On Error GoTo Err_Result
   
Dim intLength As Double, strArticle As String

If IsMissing(varKeepArticle) Then
   varKeepArticle = False
End If
intLength = Len(strOldTitle)
strArticle = ""
 
' Check Value for preceding article ("a", "an", or "the").
If Left(strOldTitle, 2) = "a " Then
   strArticle = ", " & Left(strOldTitle, 1)
   strOldTitle = Right(strOldTitle, intLength - 2)

ElseIf Left(strOldTitle, 3) = "an " Then
   strArticle = ", " & Left(strOldTitle, 2)
   strOldTitle = Right(strOldTitle, intLength - 3)
ElseIf Left(strOldTitle, 4) = "the " Then
   strArticle = ", " & Left(strOldTitle, 3)
   strOldTitle = Right(strOldTitle, intLength - 4)
End If
   
' If varKeepArticle is TRUE, then add the article string to the end.
If varKeepArticle Then
   ParseArticle = strOldTitle & strArticle
Else
   ParseArticle = strOldTitle
End If

Exit Function

Err_Result:
  ParseArticle = "#Error"

End Function

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Ïðèìåð
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Sub chtenie()
Dim txt
Dim alltxt
Dim mystring, mynumber
Open "c:\vs.txt" For Input As #1

Do Until EOF(1)
   Line Input #1, txt
'  alltxt = alltxt + txt + vbCrLf
   'Debug.Print GetWord(txt, 1)
Debug.Print Mid(txt, 1, 10)
Loop

'Do While Not EOF(1)    ' Loop until end of file.
'    Input #1, mystring, mynumber    ' Read data into two variables.
'    Debug.Print mystring, mynumber    ' Print data to the Immediate window.
'Loop

Close #1
Debug.Print alltxt
End Sub

'Ôóíêöèÿ ïðèðèòåòàò äëÿ Àíäðþêîâà
Function Poisk_str(StrSql As String)
Select Case StrSql
       Case StrSql Like "Polish": Poisk_str = "1"
       Case StrSql Like "Paint": Poisk_str = "2"
       Case StrSql Like "Replac": Poisk_str = "3"
       Case StrSql Like "Missing": Poisk_str = "4"
       Case StrSql Like "Recall": Poisk_str = "5"
       Case Else: Poisk_str = "ss"
End Select
End Function
