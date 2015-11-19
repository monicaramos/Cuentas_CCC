VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarCodempre As String 'codigo empresa
Private mvarNomempre As String 'nombre empresa

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

'Si el usuario ha modificado los datos de la empresa
'para saber si hay que actualizar los datos empresa visibles
'en el form frmPpal
Private mvarModificado As Boolean





'Private mvarClaEmpre As Byte 'clase de empresa (1=Minorista, 2=Mayorista)
'Private mvarCodGrupo As Integer 'Grupo de Empresa


Private mvarConta As String 'BD de la Contabilidad de la empresa


'Datos de la contabilidad
'--------------------------------------------
Private mvarNumnivel As Integer 'copia local
Private mvarNumdigi1 As Integer 'copia local
Private mvarNumdigi2 As Integer 'copia local
Private mvarNumdigi3 As Integer 'copia local
Private mvarNumdigi4 As Integer 'copia local
Private mvarNumdigi5 As Integer 'copia local
Private mvarNumdigi6 As Integer 'copia local
Private mvarNumdigi7 As Integer 'copia local
Private mvarNumdigi8 As Integer 'copia local
Private mvarNumdigi9 As Integer 'copia local
Private mvarNumdigi10 As Integer 'copia local

'Campo calculado de numero de digitos en la ULTIMA NIVEL CONTABLE
Private mvarDigitosUltimoNivel As Byte



'------------------------------------------------
'Propiedades del modulo Cempresa
'------------------------------------------------

'Codigo de la empresa del usuario q hizo login
Public Property Let codEmpre(ByVal vData As String)
     mvarCodempre = vData
End Property

Public Property Get codEmpre() As String
     codEmpre = mvarCodempre
End Property


'nombre de la empresa conectada
Public Property Let nomEmpre(ByVal vData As String)
     mvarNomempre = vData
End Property

Public Property Get nomEmpre() As String
     nomEmpre = mvarNomempre
End Property



Public Property Let DomicilioEmpresa(ByVal vData As String)
     mvarDomempre = vData
End Property

Public Property Get DomicilioEmpresa() As String
     DomicilioEmpresa = mvarDomempre
End Property


Public Property Let CPostal(ByVal vData As String)
     mvarCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mvarCodpobla
End Property


Public Property Let Poblacion(ByVal vData As String)
     mvarPobempre = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mvarPobempre
End Property

Public Property Let Provincia(ByVal vData As String)
     mvarProempre = vData
End Property

Public Property Get Provincia() As String
     Provincia = mvarProempre
End Property

Public Property Let CifEmpresa(ByVal vData As String)
     mvarCifempre = vData
End Property

Public Property Get CifEmpresa() As String
     CifEmpresa = mvarCifempre
End Property

Public Property Let Telefono(ByVal vData As String)
     mvarTelempre = vData
End Property

Public Property Get Telefono() As String
     Telefono = mvarTelempre
End Property

Public Property Let Fax(ByVal vData As String)
     mvarFaxempre = vData
End Property

Public Property Get Fax() As String
     Fax = mvarFaxempre
End Property


Public Property Let WebEmpresa(ByVal vData As String)
     mvarWwwempre = vData
End Property

Public Property Get WebEmpresa() As String
     WebEmpresa = mvarWwwempre
End Property

Public Property Let MailEmpresa(ByVal vData As String)
     mvarMaiempre = vData
End Property

Public Property Get MailEmpresa() As String
     MailEmpresa = mvarMaiempre
End Property

Public Property Let PerEmpresa(ByVal vData As String)
     mvarPerempre = vData
End Property

Public Property Get PerEmpresa() As String
     PerEmpresa = mvarPerempre
End Property

Public Property Let Modificado(ByVal vData As Boolean)
     mvarModificado = vData
End Property

Public Property Get Modificado() As Boolean
     Modificado = mvarModificado
End Property


'Datos de la Contabilidad
'-------------------------------------------------

'base de datos de la contabilidad de la empresa
Public Property Let BDConta(ByVal vData As String)
     mvarConta = vData
End Property

Public Property Get BDConta() As String
     BDConta = mvarConta
End Property


Public Property Let numNivel(ByVal vData As Integer)
     mvarNumnivel = vData
End Property

Public Property Get numNivel() As Integer
     numNivel = mvarNumnivel
End Property


Public Property Let numDigi1(ByVal vData As Integer)
     mvarNumdigi1 = vData
End Property

Public Property Get numDigi1() As Integer
     numDigi1 = mvarNumdigi1
End Property


Public Property Let numDigi2(ByVal vData As Integer)
     mvarNumdigi2 = vData
End Property

Public Property Get numDigi2() As Integer
     numDigi2 = mvarNumdigi2
End Property


Public Property Let numDigi3(ByVal vData As Integer)
     mvarNumdigi3 = vData
End Property

Public Property Get numDigi3() As Integer
     numDigi3 = mvarNumdigi3
End Property


Public Property Let numDigi4(ByVal vData As Integer)
     mvarNumdigi4 = vData
End Property

Public Property Get numDigi4() As Integer
     numDigi4 = mvarNumdigi4
End Property


Public Property Let numDigi5(ByVal vData As Integer)
     mvarNumdigi5 = vData
End Property

Public Property Get numDigi5() As Integer
     numDigi5 = mvarNumdigi5
End Property

Public Property Let numDigi6(ByVal vData As Integer)
     mvarNumdigi6 = vData
End Property

Public Property Get numDigi6() As Integer
     numDigi6 = mvarNumdigi6
End Property


Public Property Let numDigi7(ByVal vData As Integer)
     mvarNumdigi7 = vData
End Property

Public Property Get numDigi7() As Integer
     numDigi7 = mvarNumdigi7
End Property

Public Property Let numDigi8(ByVal vData As Integer)
     mvarNumdigi8 = vData
End Property

Public Property Get numDigi8() As Integer
     numDigi8 = mvarNumdigi8
End Property


Public Property Let numDigi9(ByVal vData As Integer)
     mvarNumdigi9 = vData
End Property

Public Property Get numDigi9() As Integer
     numDigi9 = mvarNumdigi9
End Property


Public Property Let numDigi10(ByVal vData As Integer)
     mvarNumdigi10 = vData
End Property

Public Property Get numDigi10() As Integer
     numDigi10 = mvarNumdigi10
End Property


Public Property Get DigitosUltimoNivel() As Byte
     DigitosUltimoNivel = mvarDigitosUltimoNivel
End Property


'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------

'****   AGREGAR  ***
'Public Function Agregar() As Byte
''Añadir una nueva empresa
'    On Error GoTo Err1
'
''    SQL = "INSERT INTO empresa (codempre, nomempre, nomresum, numnivel, numdigi1,"
''    SQL = SQL & "numdigi2, numdigi3, numdigi4, numdigi5, numdigi6, numdigi7,"
''    SQL = SQL & " numdigi8, numdigi9, numdigi10) VALUES ("
'
'    SQL = "INSERT INTO empresa (codempre, nomempre, nomresum"
'    SQL = SQL & " VALUES ("
'
'
'    'Datos
'    SQL = SQL & mvarCodempre
'    SQL = SQL & ",'" & mvarNomempre
'    SQL = SQL & "','" & mvarNomresum
''    SQL = SQL & "'," & mvarnumnivel
''    SQL = SQL & "," & mvarnumdigi1
''    SQL = SQL & "," & mvarnumdigi2
''    SQL = SQL & "," & mvarnumdigi3
''    SQL = SQL & "," & mvarnumdigi4
''    SQL = SQL & "," & mvarnumdigi5
''    SQL = SQL & "," & mvarnumdigi6
''    SQL = SQL & "," & mvarnumdigi7
''    SQL = SQL & "," & mvarnumdigi8
''    SQL = SQL & "," & mvarnumdigi9
''    SQL = SQL & "," & mvarnumdigi10 & ")"
'     SQL = SQL & "," & ")"
'
'    Conn.Execute SQL
'
'    Exit Function
'Err1:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    RS.Close
'    Agregar = 1
'End Function



'****   LEER  ***
Public Function LeerDatos(Codigo As Integer) As Boolean
'Leer los parametros de una empresa
'Lee de la BD: PlannerTours, Tabla: empresas
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM sempre WHERE codempre=1" ' & codigo 'vUsu.CadenaConexion


    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        LeerDatos = False
    Else
'        Mens = "Datos basicos"
        mvarCodempre = RS!codEmpre
        mvarNomempre = RS!nomEmpre
        mvarDomempre = RS!domempre
        mvarCodpobla = RS!codPosta
        mvarPobempre = RS!pobempre
        mvarProempre = RS!proempre
        mvarCifempre = RS!nifempre
        mvarTelempre = DBLet(RS!telempre, "T")
        mvarFaxempre = DBLet(RS!faxempre, "T")
        mvarWwwempre = DBLet(RS!wwwempre, "T")
        mvarMaiempre = DBLet(RS!maiempre, "T")
        mvarPerempre = DBLet(RS!perempre, "T")
        
        mvarModificado = False
'        mvarCodGrupo = RS!codGrupo
'        mvarClaEmpre = RS!claEmpre
        
        mvarConta = "conta2" ' DBLet(RS!bdaconta)
        
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Leyendo datos de la empresa." & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function




'****   MODIFICAR  ***
Public Function ModificarDatos() As Byte
'Modifica los parametros de una empresa
'Modifica la Tabla: empresasariges, BD: Usuarios
    On Error GoTo Err4

'Modificamos en la tabla Usuarios.empresasariges
'    SQL = "UPDATE Usuarios.empresasariges SET nomempre= '" & mvarNomempre
'    SQL = SQL & "', nomresum= '" & mvarNomresum
'    SQL = SQL & "' WHERE codempre=" & mvarCodempre
'    Conn.Execute SQL

    Exit Function

Err4:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    RS.Close
'    Modificar = 1
End Function

