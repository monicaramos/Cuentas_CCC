VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros: SPARAM
Private mvarCodigo As Integer       'Codigo tabla parametros empresa
Private mvarNomempre As String      'Nombre Empresa
Private mvarDomempre As String      'Domicilio Empresa
Private mvarCodpobla As String      'Cod. Postal empresa
Private mvarPobempre As String      'Poblacion empresa
Private mvarProempre As String      'Provincia empresa
Private mvarCifempre As String * 9  'C.I.F. empresa
Private mvarTelempre As String      'Teléfono empresa
Private mvarFaxempre As String      'Fax empresa
Private mvarWwwempre As String      'web empresa
Private mvarMaiempre As String      'e-mail empresa
Private mvarPerempre As String      'Gerente
'
''Si el usuario ha modificado los datos de la empresa
''para saber si hay que actualizar los datos empresa visibles
''en el form frmPpal
Private mvarModificado As Boolean


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As Recordset
Private RC As Byte



'------------------------------------------------
'Propiedades del modulo Cparametros
'------------------------------------------------

Public Property Let Codigo(ByVal vData As Integer)
     mvarCodigo = vData
End Property

Public Property Get Codigo() As Integer
     Codigo = mvarCodigo
End Property


Public Property Let NombreEmpresa(ByVal vData As String)
     mvarNomempre = vData
End Property

Public Property Get NombreEmpresa() As String
     NombreEmpresa = mvarNomempre
End Property




'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
    On Error GoTo Err3
    SQL = "SELECT * FROM sempre "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        Leer = 1
    Else
        mvarNomempre = Rs!nomEmpre
        mvarDomempre = Rs!domempre
        mvarCodpobla = Rs!codPobla
        mvarPobempre = Rs!pobempre
        mvarProempre = Rs!proempre
        mvarCifempre = Rs!cifempre
        mvarTelempre = DBLet(Rs!telempre, "T")
        mvarFaxempre = DBLet(Rs!faxempre, "T")
        mvarWwwempre = DBLet(Rs!wwwempre, "T")
        mvarMaiempre = DBLet(Rs!maiempre, "T")
        mvarPerempre = DBLet(Rs!perempre, "T")
        
        mvarModificado = False
        Leer = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Leer = 1
End Function


Public Function Modificar() As Byte
'Modifica los parametros de una empresa
'Modifica la Tabla: sparam, BD: Ariges
    On Error GoTo Err4

    mvarModificado = True

    SQL = "UPDATE sparam SET "
    SQL = SQL & " nomempre = " & DBSet(mvarNomempre, "T")
    SQL = SQL & ", domempre = " & DBSet(mvarDomempre, "T")
    SQL = SQL & ", codpobla = '" & mvarCodpobla
    SQL = SQL & "', pobempre = " & DBSet(mvarPobempre, "T")
    SQL = SQL & ", proempre = " & DBSet(mvarProempre, "T")
    SQL = SQL & ", cifempre = '" & mvarCifempre
    SQL = SQL & "', telempre = '" & mvarTelempre
    SQL = SQL & "', faxempre = '" & mvarFaxempre
    SQL = SQL & "', wwwempre = " & DBSet(mvarWwwempre, "T")
    SQL = SQL & ", maiempre = " & DBSet(mvarMaiempre, "T")
    SQL = SQL & ", perempre = " & DBSet(mvarPerempre, "T")
    SQL = SQL & " WHERE codigo =1" & ";"

    Conn.Execute SQL

    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Modificar = 1
End Function

