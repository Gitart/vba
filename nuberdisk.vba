Attribute VB_Name = "Module3"
Option Compare Database
Option Explicit

Function NomDisk()
Dim fso, objDrive, s
Set fso = CreateObject("Scripting.FileSystemObject")
Set objDrive = fso.GetDrive("f:")
NomDisk = Abs(objDrive.SerialNumber)

End Function
