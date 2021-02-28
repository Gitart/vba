Option Compare Database
Option Explicit
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Çàêà÷êà äàííûõ èç CSV Vista
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Sub Zakachka()
 DoCmd.RunSQL "DELETE * FROM EXPORTALL"
 DoCmd.TransferText acImportDelim, "EXPDS", "EXPORTALL", "c:\WINNER\DS\CSV\EXPORT.CSV", True
 DoCmd.RunSQL "UPDATE EXPORTALL SET Brend=4 WHERE Brend=0 OR Brend IS NULL"
End Sub

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Çàêà÷êà
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Sub zapolnenie()
Dim Rk As Recordset
Dim Norder As Double
Dim Stroka As String
Dim Strokad As String

Set Rk = CurrentDb().OpenRecordset("EXPORTALL")

Do While Not Rk.EOF
   Norder = Nz(DLookup("Id", "W_Order", "Zakaz='" & Rk![Order no] & "'"), 0)
   
Stroka = Rk!Ïîëå108 & " " & Rk!Ïîëå109 & Chr(10) _
       & Rk!Ïîëå110 & " " & Rk!Ïîëå111 & Chr(10) _
       & Rk!Ïîëå112 & " " & Rk!Ïîëå113 & Chr(10) _
       & Rk!Ïîëå114 & " " & Rk!Ïîëå115 & Chr(10) _
       & Rk!Ïîëå116 & " " & Rk!ÏÎËÅ117 & Chr(10) _
       & Rk!Ïîëå118 & " " & Rk!Ïîëå119 & Chr(10) _
       & Rk!Ïîëå120 & " " & Rk!Ïîëå121 & Chr(10) _
       & Rk!Ïîëå122 & " " & Rk!ÏÎËÅ123 & Chr(10) _
       & Rk!Ïîëå124 & " " & Rk!Ïîëå125 & Chr(10) _
       & Rk!ÏÎËÅ126 & " " & Rk!ÏÎËÅ127 & Chr(10) _
       & Rk!Ïîëå128 & " " & Rk!Ïîëå129 & Chr(10) _
       & Rk!ÏÎËÅ130 & " " & Rk!Ïîëå131 & Chr(10) _
       & Rk!ÏÎËÅ132 & " " & Rk!ÏÎËÅ133 & Chr(10) _
       & Rk!Ïîëå134 & " " & Rk!Ïîëå135 & Chr(10) _
       & Rk!ÏÎËÅ136 & " " & Rk!Ïîëå137 & Chr(10) _
       & Rk!Ïîëå138 & " " & Rk!Ïîëå139
         
        'Óäàëåíèå èçëèøåñòâ íåõîðîøèõ
         Stroka = Replace(Stroka, "'", " ")
         Stroka = Replace(Stroka, ",", " ")
         Stroka = Replace(Stroka, "&", " ")
         Stroka = Replace(Stroka, "END OF LINE", "")
         Stroka = Replace(Stroka, "/", "_")
         Stroka = Replace(Stroka, "(", " ")
         Stroka = Replace(Stroka, ")", " ")
    
         Strokad = Replace(Rk![DERIVATIVE DESCRIPTION], ",", " ")
         Strokad = Replace(Strokad, "/", " ")
   
'Åñëè íåò òàêîãî çàêàçà - âñòàâëÿåì íîâûé èíà÷å îáíîâëÿåì îïöèè
   If Norder = 0 Then
     'Debug.Print "äîáàâèëè = " & Rk.Fields(7)
   
      DoCmd.RunSQL "INSERT INTO W_ORDER (BREND, ZAKAZ, Sorder, DERIVATE, OPTIONS, D_Zakaza) " _
                & "VALUES (" & Rk!Brend & ",'" & Rk![Order no] & "', '" _
                              & Rk![FACTORY ORDER NUMBER] & "', '" & Strokad & "', '" & Stroka & "', #" & Rk![Order created date] & "# )"
   Else
      'Debug.Print "Îáíîâëåíèå = " & Rk.Fields(7)
      DoCmd.RunSQL "UPDATE W_ORDER SET OPTIONS = '" & Stroka & "' WHERE Zakaz='" & Rk![FACTORY ORDER NUMBER] & "'"
   End If
      Rk.MoveNext
Loop
MsgBox "Yes!!!"
End Sub

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Ïåðåíîñ äàííûõ â òàáëèöó
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Sub Pump_all()
DoCmd.RunSQL "INSERT INTO Exportall SELECT Export.* FROM Export"
End Sub

