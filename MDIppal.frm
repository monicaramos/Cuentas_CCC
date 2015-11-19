VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIppal 
   BackColor       =   &H8000000C&
   Caption         =   "Copia Versiones"
   ClientHeight    =   7860
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   11160
   Icon            =   "MDIppal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Aplicaciones"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pcs"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ficheros Aplicación"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "PCs Copia"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnParametros 
      Caption         =   "&Datos Básicos"
      Index           =   1
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Aplicaciones"
         Index           =   1
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&PCs"
         Index           =   2
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Ficheros Aplicacion"
         Index           =   3
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&PCs Copia"
         Index           =   4
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Salir"
         Index           =   6
      End
   End
End
Attribute VB_Name = "MDIppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private PrimeraVez As Boolean
Dim TieneEditorDeMenus As Boolean


'Imports System.Diagnostics
'
Dim A As FileSystemObject
'
'Private Function FileVersion(ByVal elPath As String) As String
'    Dim fvi As System.Diagnostics.FileVersionInfo
'    'Return fvi.FileVersion
'    'Return fvi.ProductVersion
'
'
'    Dim sb As New System.Text.StringBuilder
'
'    sb.AppendLine ("ProductName:     " & fvi.ProductName)
'    sb.AppendLine ("FileDescription: " & fvi.FileDescription)
'    sb.AppendLine ("FileVersion:     " & fvi.FileVersion)
'    sb.AppendLine ("ProductVersion:  " & fvi.ProductVersion)
'    sb.AppendLine ("LegalCopyright:  " & fvi.LegalCopyright)
'
'    FileVersion = sb
'End Function
'

'FileVersionInfo



'Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
'    Dim i As Integer
'    Dim tRes As ResType, iCount As Integer
'
'    opcio = op
'    tamany = tam
'    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)
'
'    If ghmodule = 0 Then
'        MsgBox "Invalid library file.", vbCritical
'        Exit Sub
'    End If
'
'    For tRes = RT_FIRST To RT_LAST
'        DoEvents
'        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
'    Next
'    FreeLibrary ghmodule
'
'End Sub

Private Sub MDIForm_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub

Private Sub MDIForm_Load()
Dim cad As String
Dim i As Integer
Dim b As String


'    Set A = New FileSystemObject
'
'    b = A.GetFileVersion("c:\programas\ariagro\ariagro.exe")
'
'
'    cad = FileVersion("c:\programas\ariagro\ariagro.exe")


    PrimeraVez = True
'    CargarImagen
'    PonerDatosFormulario

'    If vEmpresa Is Nothing Then
'        Caption = "COPIAVERSIONES" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
'    Else
        Caption = "ARICOPIA - Copia de Aplicaciones Ariadna " & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & _
                  "   -  Usuario: " & vSesion.Nombre
'    End If

    ' *** per als iconos XP ***
    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 32

    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 24
    GetIconsFromLibrary App.Path & "\iconos_BN.dll", 2, 24
    GetIconsFromLibrary App.Path & "\iconos_OM.dll", 3, 24

    GetIconsFromLibrary App.Path & "\iconosAricopia.dll", 4, 24



    'CARGAR LA TOOLBAR DEL FORM PRINCIPAL
    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListPpal

        .Buttons(1).Image = 1  'Aplicaciones
        .Buttons(2).Image = 2  'Pcs
        .Buttons(3).Image = 3  'Ficheros de aplicacion
        .Buttons(4).Image = 4  'Pcs Copia
        
    End With

    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 16
    GetIconsFromLibrary App.Path & "\iconos_BN.dll", 2, 16
    GetIconsFromLibrary App.Path & "\iconos_OM.dll", 3, 16

'    LeerEditorMenus
'
'    PonerDatosFormulario
'
'    BloqueoDeMenus

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    AccionesCerrar
    End
End Sub

Private Sub mnE_Soporte1_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome "websoporte"
    espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnE_Util_Click(Index As Integer)
    SubmnE_Util_Click (Index)
End Sub

Private Sub mnP_Generales_Click(Index As Integer)
    SubmnP_Generales_Click (Index)
End Sub



Private Sub mnP_Salir1_Click()
'    Unload frmPpal
'    Unload Me
    BotonSalir
End Sub

Private Sub mnP_Salir2_Click()
'    Unload frmPpal
'    Unload Me
    BotonSalir
End Sub

Private Sub BotonSalir()
    Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Aplicaciones
            SubmnP_Generales_Click (1)
        Case 2 'Pcs
            SubmnP_Generales_Click (2)
        Case 3 'Ficheros de Aplicacion
            SubmnP_Generales_Click (3)
        Case 4 'Pcs copia
            SubmnP_Generales_Click (4)
    End Select
End Sub

' ### [Monica] 05/09/2006
Private Sub PonerDatosFormulario()
Dim Config As Boolean

    Config = (vEmpresa Is Nothing) Or (vParamAplic Is Nothing)
    
    If Not Config Then HabilitarSoloPrametros_o_Empresas True

    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
'    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
'    PonerMenusNivelUsuario

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If Config Then HabilitarSoloPrametros_o_Empresas False
    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If


    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor

End Sub

' ### [Monica] 05/09/2006
Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim cad As String

    On Error Resume Next
    For Each T In Me
        cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            'If LCase(Mid(T.Name, 1, 8)) <> "mn_b" Then
                T.Enabled = Habilitar
            'End If
        End If
    Next
    
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnParametros(1).Enabled = True
    Me.mnP_Generales(1).Enabled = True
    Me.mnP_Generales(2).Enabled = True
    Me.mnP_Generales(6).Enabled = True
    Me.mnP_Generales(17).Enabled = True
    
'    Me.mnCambioEmpresa.Enabled = True
End Sub


' ### [Monica] 07/11/2006
' añadida esta parte para la personalizacion de menus

Private Sub LeerEditorMenus()
Dim Sql As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    Sql = "Select count(*) from appmenus where aplicacion='Avnics'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim Sql As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    Sql = "Select * from appmenususuario where aplicacion='Avnics' and codusu = " & Val(Right(CStr(vSesion.Codusu), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            Sql = Sql & miRsAux.Fields(3) & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If Sql <> "" Then
        Sql = "·" & Sql
        For Each T In Me.Controls
            If TypeOf T Is Menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                If InStr(1, Sql, C) > 0 Then T.visible = False
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index '& "|"   Monica:con esto no funcionaba
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function

Private Sub LanzaHome(Opcion As String)
    Dim i As Integer
    Dim cad As String
    On Error GoTo ELanzaHome
    
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBD("websoporte", "sparam", "codparam", 1, "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en parametros.", vbExclamation
        Exit Sub
    End If
        
    i = FreeFile
    cad = ""
    Open App.Path & "\lanzaexp.dat" For Input As #i
    Line Input #i, cad
    Close #i
    
    'Lanzamos
    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub

Private Sub CargarImagen()

On Error GoTo eCargarImagen
    Me.Picture = LoadPicture(App.Path & "\fondo.dat")
    Exit Sub
eCargarImagen:
    MuestraError Err.Number, "Error cargando imagen. LLame a soporte"
    End
End Sub

Private Sub BloqueoDeMenus()
Dim b As Boolean
'    mnAvnics.visible = (vParamAplic.Avnics = 1)
'    mnAvnics.Enabled = (vParamAplic.Avnics = 1)
'
'    mnSeguros.visible = (vParamAplic.Seguros = 1)
'    mnSeguros.Enabled = (vParamAplic.Seguros = 1)
'
'    mnTelefonia.visible = (vParamAplic.Telefonia = 1)
'    mnTelefonia.Enabled = (vParamAplic.Telefonia = 1)
End Sub

Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub

