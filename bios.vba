Осуществить привязку программы проще всего к дате создания BIOS материнской платы. Адрес расположения даты в памяти: F000:FFF5. Чтобы считать дату из BIOS, воспользуйтесь нижеследующим кодом:
Dmitry Sergunin: 

   Type BIOS_DATE
      s As String * 8
   End Type

   Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
           (pDest As Any, pSource As Any, ByVal ByteLen As Long)

   Public Function BIOS() As Long
      Dim sDB As BIOS_DATE

      CopyMemory sDB, ByVal &HFFFF5, 8&
      BIOS = DateSerial(Mid(sDB.s, 7, 2), Mid(sDB.s, 1, 2), Mid(sDB.s, 4, 2))
   End Function

 

