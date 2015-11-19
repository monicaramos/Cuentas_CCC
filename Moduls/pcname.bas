Attribute VB_Name = "pcname"
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Public Const MAX_COMPUTERNAME_LENGTH = 255


Public Function ComputerName() As String
    'Devuelve el nombre del equipo actual
    Dim sComputerName As String
    Dim ComputerNameLength As Long
    
    sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    ComputerNameLength = MAX_COMPUTERNAME_LENGTH
    Call GetComputerName(sComputerName, ComputerNameLength)
     ComputerName = Mid(sComputerName, 1, ComputerNameLength)
    
End Function


