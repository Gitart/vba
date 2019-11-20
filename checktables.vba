Возвращает true если есть таблица в базе
'******************** Code Start ************************
Function НаличиеТаблицыБазы(Str As String) As Boolean
'Возвращает true если есть таблица в базе
  НаличиеТаблицыБазы = False
  On Error GoTo Met1
  Dim SearchString As String
  SearchString = CurrentDb.TableDefs(Str).Connect
  НаличиеТаблицыБазы = True
Exit Function
Met1:
  НаличиеТаблицыБазы = False
End Function
'******************** Code End ************************
