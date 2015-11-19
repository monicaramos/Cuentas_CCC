VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPARAM

'Datos de la Contabilidad
Private mvarSerConta As String 'Servidor BD de Contabilidad
Private mvarUsuConta As String 'Usuario BD de Contabilidad
Private mvarPasConta As String 'Password BD de Contabilidad
Private mvarNumConta As Byte 'Numero BD de Contabilidad
Private mvarCtaGasto As String 'Cuenta Gastos
Private mvarCtaReten As String 'Cuenta Retencion
Private mvarPorcinte As String
Private mvarPorcrete As String
Private mvarConceDebe As String 'Concepto al debe
Private mvarConceHaber As String 'Concepto al haber
Private mvarNumDiario As String 'Numero de Diario
Private mvarSeguros As String 'hay seguros
Private mvarAvnics As String 'hay avnics
Private mvarTelefonia As String 'hay telefonia

' para seguros
Private mvarSerContaSeg As String 'Servidor BD de Contabilidad
Private mvarUsuContaSeg As String 'Usuario BD de Contabilidad
Private mvarPasContaSeg As String 'Password BD de Contabilidad
Private mvarNumContaSeg As Byte 'Numero BD de Contabilidad
Private mvarCtaBancoSeg As String 'Cuenta contable de banco
Private mvarConceDebeSeg As String 'Concepto al debe
Private mvarConceHaberSeg As String 'Concepto al haber
Private mvarNumDiarioSeg As String 'Numero de Diario
Private mvarRaizCtaSocSeg As String 'raiz cta contable socio

' para telefonia
Private mvarSerContaTel As String 'Servidor BD de Contabilidad
Private mvarUsuContaTel As String 'Usuario BD de Contabilidad
Private mvarPasContaTel As String 'Password BD de Contabilidad
Private mvarNumContaTel As Byte 'Numero BD de Contabilidad
Private mvarConceDebeTel As String 'Concepto al debe
Private mvarConceHaberTel As String 'Concepto al haber
Private mvarNumDiarioTel As String 'Numero de Diario
Private mvarRaizCtaSocTel As String 'raiz cta contable socio
Private mvarCtaBancoTel As String 'Cuenta contable de banco
Private mvarCtaVentaTel As String 'Cuenta contable de venta
Private mvarNumSerieTel As String 'Numero de Serie de Telefonia

' para factura varias
Private mvarFacturasVarias As Byte 'indicador de si hay facturas varias
Private mvarSerContaFac As String 'Servidor BD de Contabilidad
Private mvarUsuContaFac As String 'Usuario BD de Contabilidad
Private mvarPasContaFac As String 'Password BD de Contabilidad

Private mvarWebSoporte As String  'Web Soporte
Private mvarDireMail As String 'direccion mail
Private mvarSmtphost As String 'Host
Private mvarSmtpUser As String 'User
Private mvarSmtpPass As String 'Password

Private Sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes


'------------------------------------------------
'Propiedades del modulo CParamAplic
'------------------------------------------------

'servidor conta
Public Property Let ServidorConta(ByVal vData As String)
     mvarSerConta = vData
End Property

Public Property Get ServidorConta() As String
     ServidorConta = mvarSerConta
End Property

Public Property Let UsuarioConta(ByVal vData As String)
     mvarUsuConta = vData
End Property

Public Property Get UsuarioConta() As String
     UsuarioConta = mvarUsuConta
End Property

Public Property Let PasswordConta(ByVal vData As String)
     mvarPasConta = vData
End Property

Public Property Get PasswordConta() As String
     PasswordConta = mvarPasConta
End Property

Public Property Let NumeroConta(ByVal vData As Byte)
     mvarNumConta = vData
End Property

Public Property Get NumeroConta() As Byte
     NumeroConta = mvarNumConta
End Property

Public Property Let CtaGasto(ByVal vData As String)
     mvarCtaGasto = vData
End Property

Public Property Get CtaGasto() As String
     CtaGasto = mvarCtaGasto
End Property

Public Property Let CtaReten(ByVal vData As String)
     mvarCtaReten = vData
End Property

Public Property Get CtaReten() As String
     CtaReten = mvarCtaReten
End Property

Public Property Let Porcinte(ByVal vData As String)
     mvarPorcinte = vData
End Property

Public Property Get Porcinte() As String
     Porcinte = mvarPorcinte
End Property

Public Property Let Porcrete(ByVal vData As String)
     mvarPorcrete = vData
End Property

Public Property Get Porcrete() As String
     Porcrete = mvarPorcrete
End Property

Public Property Let ConceDebe(ByVal vData As Long)
     mvarConceDebe = vData
End Property

Public Property Get ConceDebe() As Long
     ConceDebe = mvarConceDebe
End Property

Public Property Let ConceHaber(ByVal vData As Long)
     mvarConceHaber = vData
End Property

Public Property Get ConceHaber() As Long
     ConceHaber = mvarConceHaber
End Property

Public Property Let NumDiario(ByVal vData As Long)
     mvarNumDiario = vData
End Property

Public Property Get NumDiario() As Long
     NumDiario = mvarNumDiario
End Property


Public Property Let WebSoporte(ByVal vData As String)
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
     WebSoporte = mvarWebSoporte
End Property

Public Property Let DireMail(ByVal vData As String)
     mvarDireMail = vData
End Property

Public Property Get DireMail() As String
     DireMail = mvarDireMail
End Property

Public Property Let Smtphost(ByVal vData As String)
     mvarSmtphost = vData
End Property

Public Property Get Smtphost() As String
     Smtphost = mvarSmtphost
End Property

Public Property Let SmtpUser(ByVal vData As String)
     mvarSmtpUser = vData
End Property

Public Property Get SmtpUser() As String
     SmtpUser = mvarSmtpUser
End Property

Public Property Let Smtppass(ByVal vData As String)
     mvarSmtpPass = vData
End Property

Public Property Get Smtppass() As String
     Smtppass = mvarSmtpPass
End Property

Public Property Let Seguros(ByVal vData As Byte)
     mvarSeguros = vData
End Property

Public Property Get Seguros() As Byte
     Seguros = mvarSeguros
End Property

Public Property Let Avnics(ByVal vData As Byte)
     mvarAvnics = vData
End Property

Public Property Get Avnics() As Byte
     Avnics = mvarAvnics
End Property


Public Property Let Telefonia(ByVal vData As Byte)
     mvarTelefonia = vData
End Property

Public Property Get Telefonia() As Byte
     Telefonia = mvarTelefonia
End Property

Public Property Let FacturasVarias(ByVal vData As Byte)
     mvarFacturasVarias = vData
End Property

Public Property Get FacturasVarias() As Byte
     FacturasVarias = mvarFacturasVarias
End Property



'servidor conta
Public Property Let ServidorContaSeg(ByVal vData As String)
     mvarSerContaSeg = vData
End Property

Public Property Get ServidorContaSeg() As String
     ServidorContaSeg = mvarSerContaSeg
End Property

Public Property Let UsuarioContaSeg(ByVal vData As String)
     mvarUsuContaSeg = vData
End Property

Public Property Get UsuarioContaSeg() As String
     UsuarioContaSeg = mvarUsuContaSeg
End Property

Public Property Let PasswordContaSeg(ByVal vData As String)
     mvarPasContaSeg = vData
End Property

Public Property Get PasswordContaSeg() As String
     PasswordContaSeg = mvarPasContaSeg
End Property

Public Property Let NumeroContaSeg(ByVal vData As Byte)
     mvarNumContaSeg = vData
End Property

Public Property Get NumeroContaSeg() As Byte
     NumeroContaSeg = mvarNumContaSeg
End Property

Public Property Let CtaBancoSeg(ByVal vData As String)
     mvarCtaBancoSeg = vData
End Property

Public Property Get CtaBancoSeg() As String
     CtaBancoSeg = mvarCtaBancoSeg
End Property

Public Property Let ConceDebeSeg(ByVal vData As Long)
     mvarConceDebeSeg = vData
End Property

Public Property Get ConceDebeSeg() As Long
     ConceDebeSeg = mvarConceDebeSeg
End Property

Public Property Let ConceHaberSeg(ByVal vData As Long)
     mvarConceHaberSeg = vData
End Property

Public Property Get ConceHaberSeg() As Long
     ConceHaberSeg = mvarConceHaberSeg
End Property

Public Property Let NumDiarioSeg(ByVal vData As Long)
     mvarNumDiarioSeg = vData
End Property

Public Property Get NumDiarioSeg() As Long
     NumDiarioSeg = mvarNumDiarioSeg
End Property

Public Property Let RaizCtaSocSeg(ByVal vData As String)
     mvarRaizCtaSocSeg = vData
End Property

Public Property Get RaizCtaSocSeg() As String
     RaizCtaSocSeg = mvarRaizCtaSocSeg
End Property


' telefonia
'servidor conta
Public Property Let ServidorContaTel(ByVal vData As String)
     mvarSerContaTel = vData
End Property

Public Property Get ServidorContaTel() As String
     ServidorContaTel = mvarSerContaTel
End Property

Public Property Let UsuarioContaTel(ByVal vData As String)
     mvarUsuContaTel = vData
End Property

Public Property Get UsuarioContaTel() As String
     UsuarioContaTel = mvarUsuContaTel
End Property

Public Property Let PasswordContaTel(ByVal vData As String)
     mvarPasContaTel = vData
End Property

Public Property Get PasswordContaTel() As String
     PasswordContaTel = mvarPasContaTel
End Property

Public Property Let NumeroContaTel(ByVal vData As Byte)
     mvarNumContaTel = vData
End Property

Public Property Get NumeroContaTel() As Byte
     NumeroContaTel = mvarNumContaTel
End Property

Public Property Let ConceDebeTel(ByVal vData As Long)
     mvarConceDebeTel = vData
End Property

Public Property Get ConceDebeTel() As Long
     ConceDebeTel = mvarConceDebeTel
End Property

Public Property Let ConceHaberTel(ByVal vData As Long)
     mvarConceHaberTel = vData
End Property

Public Property Get ConceHaberTel() As Long
     ConceHaberTel = mvarConceHaberTel
End Property

Public Property Let NumDiarioTel(ByVal vData As Long)
     mvarNumDiarioTel = vData
End Property

Public Property Get NumDiarioTel() As Long
     NumDiarioTel = mvarNumDiarioTel
End Property

Public Property Let RaizCtaSocTel(ByVal vData As String)
     mvarRaizCtaSocTel = vData
End Property

Public Property Get RaizCtaSocTel() As String
     RaizCtaSocTel = mvarRaizCtaSocTel
End Property

Public Property Let CtaBancoTel(ByVal vData As String)
     mvarCtaBancoTel = vData
End Property

Public Property Get CtaBancoTel() As String
     CtaBancoTel = mvarCtaBancoTel
End Property

Public Property Let CtaVentaTel(ByVal vData As String)
     mvarCtaVentaTel = vData
End Property

Public Property Get CtaVentaTel() As String
     CtaVentaTel = mvarCtaVentaTel
End Property

Public Property Let NumSerieTel(ByVal vData As String)
     mvarNumSerieTel = vData
End Property

Public Property Get NumSerieTel() As String
     NumSerieTel = mvarNumSerieTel
End Property

'facturas varias
Public Property Let ServidorContaFac(ByVal vData As String)
     mvarSerContaFac = vData
End Property

Public Property Get ServidorContaFac() As String
     ServidorContaFac = mvarSerContaFac
End Property

Public Property Let UsuarioContaFac(ByVal vData As String)
     mvarUsuContaFac = vData
End Property

Public Property Get UsuarioContaFac() As String
     UsuarioContaFac = mvarUsuContaFac
End Property

Public Property Let PasswordContaFac(ByVal vData As String)
     mvarPasContaFac = vData
End Property

Public Property Get PasswordContaFac() As String
     PasswordContaFac = mvarPasContaFac
End Property


'******
'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function leer() As Byte
Dim RS As ADODB.Recordset
On Error GoTo Err3
    
    Sql = "SELECT * FROM sparam "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        leer = 1
    Else
        'Datos Contabilidad avnics
        mvarSerConta = DBLet(RS!serConta, "T") 'servidor conta
        mvarUsuConta = DBLet(RS!UsuConta, "T") 'usuario conta
        mvarPasConta = DBLet(RS!pasconta, "T") 'password
        mvarNumConta = DBLet(RS!NumConta, "N") 'numero conta
        mvarCtaGasto = DBLet(RS!CtaGasto, "T")
        mvarCtaReten = DBLet(RS!CtaReten, "T")
        mvarPorcinte = DBLet(RS!Porcinte, "N")
        mvarPorcrete = DBLet(RS!Porcrete, "N")
        mvarConceDebe = DBLet(RS!ConceDebe, "N") 'concepto al debe
        mvarConceHaber = DBLet(RS!ConceHaber, "N") 'concepto al haber
        mvarNumDiario = DBLet(RS!NumDiario, "N") 'numero de diario
        
        mvarWebSoporte = DBLet(RS!WebSoporte, "T") 'web soporte
        mvarDireMail = DBLet(RS!DireMail, "T") 'mail
        mvarSmtphost = DBLet(RS!Smtphost, "T") 'host
        mvarSmtpUser = DBLet(RS!SmtpUser, "T") 'user
        mvarSmtpPass = DBLet(RS!Smtppass, "T") 'pass
        
        mvarSeguros = DBLet(RS!Seguros, "N") 'hay seguros agrarios
        mvarAvnics = DBLet(RS!Avnics, "N") 'hay gestion de avnics
        mvarTelefonia = DBLet(RS!Telefonia, "N") 'hay telefonia
        mvarFacturasVarias = DBLet(RS!FacturasVarias, "N") 'hay facturasvarias
        
        'datos contabilidad para seguros
        mvarSerContaSeg = DBLet(RS!serContaseg, "T") 'servidor conta
        mvarUsuContaSeg = DBLet(RS!UsuContaseg, "T") 'usuario conta
        mvarPasContaSeg = DBLet(RS!pascontaseg, "T") 'password
        mvarNumContaSeg = DBLet(RS!NumContaseg, "N") 'numero conta
        mvarConceDebeSeg = DBLet(RS!ConceDebeSeg, "N") 'concepto al debe
        mvarConceHaberSeg = DBLet(RS!ConceHaberSeg, "N") 'concepto al haber
        mvarNumDiarioSeg = DBLet(RS!NumDiarioSeg, "N") 'numero de diario
        mvarCtaBancoSeg = DBLet(RS!CtaBancoSeg, "T") 'cta contable banco seguros
        mvarRaizCtaSocSeg = DBLet(RS!RaizCtaSocSeg, "T") 'raiz cta contable socio seguros
        
        'datos contabilidad para telefonia
        mvarSerContaTel = DBLet(RS!serContatel, "T") 'servidor conta
        mvarUsuContaTel = DBLet(RS!UsuContatel, "T") 'usuario conta
        mvarPasContaTel = DBLet(RS!pascontatel, "T") 'password
        mvarNumContaTel = DBLet(RS!NumContatel, "N") 'numero conta
        mvarConceDebeTel = DBLet(RS!ConceDebeTel, "N") 'concepto al debe
        mvarConceHaberTel = DBLet(RS!ConceHaberTel, "N") 'concepto al haber
        mvarNumDiarioTel = DBLet(RS!NumDiarioTel, "N") 'numero de diario
        mvarCtaBancoTel = DBLet(RS!CtaBancoTel, "T") 'cta contable banco telefonia
        mvarRaizCtaSocTel = DBLet(RS!RaizCtaSocTel, "T") 'raiz cta contable socio telefonia
        mvarCtaVentaTel = DBLet(RS!CtaVentaTel, "T") 'cta contable venta telefonia
        mvarNumSerieTel = DBLet(RS!NumSerieTel, "T") 'Numero de Serie de telefonia
        
        'datos contabilidad para facturas varias
        mvarSerContaFac = DBLet(RS!serContafac, "T") 'servidor conta
        mvarUsuContaFac = DBLet(RS!UsuContafac, "T") 'usuario conta
        mvarPasContaFac = DBLet(RS!pascontafac, "T") 'password
        leer = 0
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Set RS = Nothing
    leer = 1
End Function


Public Function Modificar(Codigo As Byte) As Byte
'Modifica los parametros de la Aplicación
'Modifica la Tabla: sparam, BD: Arigasol
On Error GoTo EModificar

    Sql = "UPDATE sparam SET "
    Sql = Sql & " serconta= '" & mvarSerConta & "'"
    Sql = Sql & ", usuconta= " & DBSet(mvarUsuConta, "T")
    Sql = Sql & ", pasconta= " & DBSet(mvarPasConta, "T")
    Sql = Sql & ", numconta= " & DBSet(mvarNumConta, "N", "S")
    Sql = Sql & ", ctagasto= " & DBSet(mvarCtaGasto, "T")
    Sql = Sql & ", ctareten = " & DBSet(mvarCtaReten, "T")
    Sql = Sql & ", porcinte= " & DBSet(mvarPorcinte, "N", "S")
    Sql = Sql & ", porcrete= " & DBSet(mvarPorcrete, "N", "S")
    Sql = Sql & ", websoporte= " & DBSet(mvarWebSoporte, "T")
    Sql = Sql & ", diremail= " & DBSet(mvarDireMail, "T")
    Sql = Sql & ", smtphost= " & DBSet(mvarSmtphost, "T")
    Sql = Sql & ", smtpuser= " & DBSet(mvarSmtpUser, "T")
    Sql = Sql & ", smtppass= " & DBSet(mvarSmtpPass, "T")
    Sql = Sql & ", concedebe= " & DBSet(mvarConceDebe, "N")
    Sql = Sql & ", concehaber= " & DBSet(mvarConceHaber, "N")
    Sql = Sql & ", numdiario= " & DBSet(mvarNumDiario, "N")
    Sql = Sql & ", seguros= " & DBSet(mvarSeguros, "N")
    Sql = Sql & ", avnics= " & DBSet(mvarAvnics, "N")
    Sql = Sql & ", telefonia= " & DBSet(mvarTelefonia, "N")
    Sql = Sql & ", facturasvarias= " & DBSet(mvarFacturasVarias, "N")
    'seguros
    Sql = Sql & ", sercontaseg= '" & mvarSerContaSeg & "'"
    Sql = Sql & ", usucontaseg= " & DBSet(mvarUsuContaSeg, "T")
    Sql = Sql & ", pascontaseg= " & DBSet(mvarPasContaSeg, "T")
    Sql = Sql & ", numcontaseg= " & DBSet(mvarNumContaSeg, "N", "S")
    Sql = Sql & ", concedebeseg= " & DBSet(mvarConceDebeSeg, "N")
    Sql = Sql & ", concehaberseg= " & DBSet(mvarConceHaberSeg, "N")
    Sql = Sql & ", numdiarioseg= " & DBSet(mvarNumDiarioSeg, "N")
    Sql = Sql & ", ctabancoseg= " & DBSet(mvarCtaBancoSeg, "T")
    Sql = Sql & ", raizctasocseg= " & DBSet(mvarRaizCtaSocSeg, "T")
    'telefonia
    Sql = Sql & ", sercontatel= '" & mvarSerContaTel & "'"
    Sql = Sql & ", usucontatel= " & DBSet(mvarUsuContaTel, "T")
    Sql = Sql & ", pascontatel= " & DBSet(mvarPasContaTel, "T")
    Sql = Sql & ", numcontatel= " & DBSet(mvarNumContaTel, "N", "S")
    Sql = Sql & ", concedebetel= " & DBSet(mvarConceDebeTel, "N")
    Sql = Sql & ", concehabertel= " & DBSet(mvarConceHaberTel, "N")
    Sql = Sql & ", numdiariotel= " & DBSet(mvarNumDiarioTel, "N")
    Sql = Sql & ", ctabancotel= " & DBSet(mvarCtaBancoTel, "T")
    Sql = Sql & ", raizctasoctel= " & DBSet(mvarRaizCtaSocTel, "T")
    Sql = Sql & ", ctaventatel= " & DBSet(mvarCtaVentaTel, "T")
    Sql = Sql & ", numserietel= " & DBSet(mvarNumSerieTel, "T")
    'facturas varias
    Sql = Sql & ", sercontafac= '" & mvarSerContaFac & "'"
    Sql = Sql & ", usucontafac= " & DBSet(mvarUsuContaFac, "T")
    Sql = Sql & ", pascontafac= " & DBSet(mvarPasContaFac, "T")
    
    Sql = Sql & " WHERE codparam =" & Codigo & ";"

    conn.Execute Sql

    Modificar = 1
    Exit Function

EModificar:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 0
End Function

