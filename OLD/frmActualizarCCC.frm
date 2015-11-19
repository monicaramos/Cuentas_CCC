VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActualizarCCC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizador de CCC"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmActualizarCCC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4050
      Width           =   2025
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3600
      Width           =   2025
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3150
      Width           =   2025
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2700
      Width           =   2025
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2250
      Width           =   2025
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   2025
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1020
      Width           =   2025
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4980
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4980
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   240
      Left            =   420
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3630
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4980
      Width           =   1275
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Columnas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   19
      Top             =   1980
      Width           =   690
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "IBAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   5
      Left            =   660
      TabIndex        =   13
      Top             =   2340
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   4
      Left            =   660
      TabIndex        =   12
      Top             =   4080
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "DC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   3
      Left            =   660
      TabIndex        =   11
      Top             =   3645
      Width           =   210
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   2
      Left            =   660
      TabIndex        =   10
      Top             =   3210
      Width           =   600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   9
      Top             =   2775
      Width           =   435
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Tabla"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   1590
      Width           =   390
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Base de datos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   11
      Left            =   360
      TabIndex        =   6
      Top             =   1020
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Actualizador de Cuentas Bancarias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   210
      Width           =   5475
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   4815
      TabIndex        =   3
      Top             =   1890
      Width           =   1950
   End
End
Attribute VB_Name = "frmActualizarCCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Aplicaciones As String
Dim NomPc As String
Dim HayPdte As Boolean
Dim cerrar As Boolean

Private Sub cmdAceptar_Click()
'Dim NomPc As String
Dim b As Boolean

    If Not DatosOk Then Exit Sub


    b = ActualizaCuentas(Combo1(1).Text, Combo1(2).Text, Combo1(3).Text, Combo1(4).Text, Combo1(5).Text, Combo1(6).Text)

    If b Then
        Label1.Caption = "Proceso realizado correctamente"
    Else
        Label1.Caption = "No se ha podido realizar el proceso correctamente. Llame a Ariadna."
    End If
    
    Me.Refresh
    DoEvents
 
    CmdSalir.SetFocus
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Function DatosOk() As Boolean
    DatosOk = Not (Combo1(0).Text = "" Or Combo1(1).Text = "" Or Combo1(2).Text = "" Or Combo1(3).Text = "" Or Combo1(4).Text = "" Or Combo1(5).Text = "" Or Combo1(6).Text = "")

End Function


Private Sub Combo1_GotFocus(Index As Integer)
    If Index = 1 Then
        If Combo1(0).Text <> "" Then
        
            LimpiarCombosCampos
        
            conn.Close
        
            AbrirConexion vConfig.User, vConfig.password, Combo1(0).Text
        
            CargaComboTabla
        End If
    
    End If
End Sub


Private Sub Combo1_LostFocus(Index As Integer)
    If Index = 1 Then
        If Combo1(Index).Text <> "" Then
            CargaComboColumnas
        End If
    End If
End Sub

Private Sub Form_Activate()
    If cerrar Then Unload Me
End Sub



Private Sub LimpiarCombosCampos()
Dim i As Integer

    For i = 2 To 6
        Combo1(i).Clear
    Next i
End Sub



Private Sub Form_Load()
Dim Cad As String, Cad1 As String
Dim Mens As String
Dim b As Boolean

    AbrirConexion vConfig.User, vConfig.password, ""

    CargaComboBD

End Sub

Private Sub CargaComboBD()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer
Dim Sql As String
Dim RS As ADODB.Recordset
Dim BDatos As Collection
    
    
    
    Combo1(0).Clear
    
    
    Set RS = New ADODB.Recordset
        
     i = 0
     
     Set BDatos = New Collection
     
     RS.Open "SHOW DATABASES", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     While Not RS.EOF
     
'        Combo1(0).AddItem CStr(RS.Fields(0)) 'campo del codigo
'        Combo1(0).ItemData(Combo1(0).NewIndex) = i
     
        If CStr(RS.Fields(0)) <> "information_schema" Then BDatos.Add CStr(RS.Fields(0))
     
        i = i + 1
     
        RS.MoveNext
     
     Wend
     RS.Close
     Set RS = Nothing
     espera 0.2
     
     For i = 1 To BDatos.Count
        Combo1(0).AddItem BDatos.Item(i) 'campo del codigo
        Combo1(0).ItemData(Combo1(0).NewIndex) = i
     Next i
     
End Sub


Private Sub CargaComboTabla()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer
Dim Sql As String
Dim RS As ADODB.Recordset
Dim LTabla As Collection
    
    
    
    Combo1(1).Clear
    
    If InStr(1, Combo1(0).Text, "ariagroutil") <> 0 Then
        LTabla.Add "avnic"
    Else
        If InStr(1, Combo1(0).Text, "ariagro") <> 0 Then
            LTabla.Add "clientes"
            LTabla.Add "proveedor"
            LTabla.Add "agencias"
            LTabla.Add "banpropi"
            LTabla.Add "rsocios"
            LTabla.Add "rtransporte"
            LTabla.Add "straba"
            LTabla.Add "clientes"
        Else
            If InStr(1, Combo1(0).Text, "conta") <> 0 Then
                LTabla.Add "cuentas"
                LTabla.Add "scobro"
                LTabla.Add "spagop"
            Else
                If InStr(1, Combo1(0).Text, "arigasol") <> 0 Then
                    LTabla.Add "ssocio"
                    LTabla.Add "starje"
                    LTabla.Add "sbanco"
                Else
                    If InStr(1, Combo1(0).Text, "ariges") <> 0 Then
                        LTabla.Add "sclien"
                        LTabla.Add "sprove"
                        LTabla.Add "straba"
                        LTabla.Add "sdirec"
                        LTabla.Add "scafac"
                        LTabla.Add "sbanpr"
                    End If
                End If
            End If
        End If
    End If
    
    For i = 1 To LTabla.Count
        Combo1(1).AddItem LTabla.Item(i)
        Combo1(1).ItemData(Combo1(1).NewIndex) = i
    Next i
    
    
'    Set RS = New ADODB.Recordset
'
'     i = 0
'     RS.Open "SHOW TABLES", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'     While Not RS.EOF
'
'        Combo1(1).AddItem CStr(RS.Fields(0)) 'campo del codigo
'        Combo1(1).ItemData(Combo1(1).NewIndex) = i
'
'        i = i + 1
'
'        RS.MoveNext
'
'     Wend
'     RS.Close
'     Set RS = Nothing
     espera 0.2
     
     
End Sub




Private Sub CargaComboColumnas()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer
Dim Sql As String
Dim RS As ADODB.Recordset
Dim L As Collection
    
    
    
    Combo1(2).Clear
    Combo1(3).Clear
    Combo1(4).Clear
    Combo1(5).Clear
    Combo1(6).Clear
    
    
    Set RS = New ADODB.Recordset
        
     i = 0
     RS.Open "SHOW COLUMNS FROM " & Combo1(1).Text, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     While Not RS.EOF
     
        Combo1(2).AddItem CStr(RS.Fields(0)) 'campo del iban
        Combo1(2).ItemData(Combo1(2).NewIndex) = i
        Combo1(3).AddItem CStr(RS.Fields(0)) 'campo del banco
        Combo1(3).ItemData(Combo1(3).NewIndex) = i
        Combo1(4).AddItem CStr(RS.Fields(0)) 'campo del sucursal
        Combo1(4).ItemData(Combo1(4).NewIndex) = i
        Combo1(5).AddItem CStr(RS.Fields(0)) 'campo del dc
        Combo1(5).ItemData(Combo1(5).NewIndex) = i
        Combo1(6).AddItem CStr(RS.Fields(0)) 'campo del cuentaba
        Combo1(6).ItemData(Combo1(6).NewIndex) = i
     
        i = i + 1
     
        RS.MoveNext
     
     Wend
     RS.Close
     Set RS = Nothing
     espera 0.2
End Sub




Private Function ActualizarVersionPC(pc As Integer, ByRef vAplic As CAplicacion) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Destino As String
Dim PathDestino As String
Dim PathFuente As String
Dim Fuente As String
Dim A As FileSystemObject
Dim NumFic As Long

Dim Fls As FileSystemObject
Dim Atributos As Integer
Dim Cambio As Boolean
Dim Nombre As String

    On Error GoTo eActualizarVersionPC

    ActualizarVersionPC = False
    
    Me.Label1.Caption = "Actualizando aplicación: " & vAplic.NomAplic
    Me.Label1.Refresh
    DoEvents
    
    Sql = "select count(*) from ficheroscopia where idaplicacion = " & DBSet(vAplic.IdAplic, "N")
    NumFic = TotalRegistros(Sql) + 1
    Me.Pb1.visible = True
    CargarProgres Me.Pb1, CInt(NumFic)
    
    
    Sql = "select * from ficheroscopia where idaplicacion = " & DBSet(vAplic.IdAplic, "N")
    
    PathDestino = ""
    PathDestino = DevuelveDesdeBDNew(cPTours, "pcscopia", "pathcopia", "idpcs", CStr(pc), "N", , "idaplicacion", CStr(vAplic.IdAplic), "N")
    
'    If InStr(1, LCase(PathDestino), "archivos de programa\ariadna") = 0 Then
'        MsgBox "El path destino es incorrecto. No se va a realizar la actualización de la aplicación " & DBLet(vAplic.NomAplic, "T"), vbExclamation
'        Exit Function
'    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not RS.EOF
        IncrementarProgres Me.Pb1, 1
  
        Fuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ)) & "\" & DBLet(RS!Nombre, "T")
        Destino = PathDestino & "\" & DBLet(RS!Nombre, "T")
        
        Set A = New FileSystemObject
        
        
        If DBLet(RS!Tipo, "N") = 0 Then ' fichero
            ' si el fichero es el ejecutable lo procesaremos el ultimo
            If InStr(1, Fuente, "exe") = 0 Then
                If InStr(1, Destino, "*") Then
                    
                    Nombre = Dir(Fuente)
                    PathFuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ))
                    Do While Len(Nombre)
                        Cambio = False
                        If A.FileExists(PathDestino & "\" & Nombre) Then
'                            Atributos = GetAttr(PathDestino & "\" & Nombre)
                            SetAttr PathDestino & "\" & Nombre, GetAttr(PathFuente & "\" & Nombre)
                            If (GetAttr(PathDestino & "\" & Nombre) And vbReadOnly) <> 0 Then
                                SetAttr PathDestino & "\" & Nombre, GetAttr(PathFuente & "\" & Nombre) - vbReadOnly 'Atributos - vbReadOnly
                                Cambio = True
                            End If
                        End If
                        FileCopy PathFuente & "\" & Nombre, PathDestino & "\" & Nombre  ' si no existiera el fichero
                        If Cambio Then
                            'SetAttr PathDestino & "\" & Nombre, GetAttr(PathDestino & "\" & Nombre) + vbReadOnly 'Atributos '+ vbReadOnly
                            SetAttr PathDestino & "\" & Nombre, GetAttr(PathFuente & "\" & Nombre)
                        End If
'                        SetAttr PathDestino & "\" & Nombre, GetAttr(PathFuente & "\" & Nombre)
                        Nombre = Dir
                    Loop
'                    If Dir(Destino) <> "" Then a.DeleteFile Destino, True
'                    a.CopyFile Fuente, PathDestino
                Else
                    Cambio = False
                    If A.FileExists(Destino) Then
'                        Atributos = GetAttr(Destino)
                        SetAttr Destino, GetAttr(Fuente)
                        If (GetAttr(Destino) And vbReadOnly) <> 0 Then
                            SetAttr Destino, GetAttr(Fuente) - vbReadOnly
                            Cambio = True
                        End If
                    End If
                    FileCopy Fuente, Destino ' si no existiera el fichero
                    
                    If Cambio Then SetAttr Destino, GetAttr(Fuente) 'SetAttr Destino, Atributos
'                    SetAttr Destino, GetAttr(Fuente)
                    'pq sino tiene problemas con el atributo de solo lectura
'                    If a.FileExists(Destino) Then a.DeleteFile Destino, True
'                    a.CopyFile Fuente, Destino
                End If
            End If
        Else
             ' carpeta
            Nombre = Dir(Fuente & "\")
            PathFuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ))
            If Not A.FolderExists(Destino) Then A.CreateFolder (Destino)
            
            Do While Len(Nombre)
                Cambio = False
                If A.FileExists(Destino & "\" & Nombre) Then
                    'Atributos = GetAttr(Destino & "\" & Nombre)
                    SetAttr Destino & "\" & Nombre, GetAttr(Fuente & "\" & Nombre)
                    If (GetAttr(Destino & "\" & Nombre) And vbReadOnly) <> 0 Then
                        SetAttr Destino & "\" & Nombre, GetAttr(Fuente & "\" & Nombre) - vbReadOnly 'Atributos - vbReadOnly
                        Cambio = True
                    End If
                End If
                FileCopy Fuente & "\" & Nombre, Destino & "\" & Nombre  ' si no existiera el fichero
                If Cambio Then
                    'SetAttr Destino & "\" & Nombre, GetAttr(Destino & "\" & Nombre) + vbReadOnly 'Atributos
                    SetAttr Destino & "\" & Nombre, GetAttr(Fuente & "\" & Nombre)
                End If
                'SetAttr Destino & "\" & Nombre, GetAttr(Fuente & "\" & Nombre)
                
                Nombre = Dir
            Loop
             
             
'            If a.FolderExists(Destino) Then a.DeleteFolder Destino, True
'            a.CopyFolder Fuente, Destino, True
        End If
            
        Set A = Nothing
    
       
    
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    'dejamos para ultimo lugar el exe
    Sql = "select * from ficheroscopia where nombre like '%exe' and idaplicacion = " & DBSet(vAplic.IdAplic, "N")
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        IncrementarProgres Me.Pb1, 1
  
        Fuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ)) & "\" & DBLet(RS!Nombre, "T")
        Destino = PathDestino & "\" & DBLet(RS!Nombre, "T")
        
        Set A = New FileSystemObject
        
        If DBLet(RS!Tipo, "N") = 0 Then ' fichero
'            'pq sino tiene problemas con el atributo de solo lectura
'            If a.FileExists(Destino) Then a.DeleteFile Destino, True
'            a.CopyFile Fuente, Destino
            Cambio = False
            If A.FileExists(Destino) Then
'                Atributos = GetAttr(Destino)
                SetAttr Destino, GetAttr(Fuente)
                If (GetAttr(Destino) And vbReadOnly) <> 0 Then
                    SetAttr Destino, GetAttr(Fuente) - vbReadOnly 'Atributos - vbReadOnly
                    Cambio = True
                End If
            End If
            
            FileCopy Fuente, Destino
        
            If Cambio Then SetAttr Destino, GetAttr(Fuente) + vbReadOnly 'Atributos
            
            'SetAttr Destino, GetAttr(Fuente)
        End If
        
    End If
    
    Set RS = Nothing
    
    Me.Pb1.visible = False

    ActualizarVersionPC = True
    Exit Function
    
eActualizarVersionPC:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Error actualizando version de " & vAplic.NomAplic & " fichero " & Destino
    End If
End Function
 


'Private Function ActualizarVersionesServidor() As Boolean
'Dim a As FileSystemObject
'
'Dim SQL As String
'Dim SQL1 As String
'Dim RS As ADODB.Recordset
'Dim Aplic As String
'Dim vAplic As CAplicacion
'Dim Version As String
'Dim Fichero As String
'Dim b As Boolean
'
'    On Error GoTo eActualizarVersionesServidor
'
'    ActualizarVersionesServidor = False
'
'    SQL = "select * from aplicaciones where idaplicacion > 0 order by idaplicacion "
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    While Not RS.EOF
'        Set vAplic = New CAplicacion
'        If vAplic.LeerDatos(RS!idaplicacion) Then
'            Label1.Caption = "Actualizando Versión: "
'            Label1.Refresh
'            DoEvents
'
'            Set a = New FileSystemObject
'
'            Fichero = "\\" & vAplic.Servidor & Mid(DBLet(RS!pathservidor, "T"), InStr(1, DBLet(RS!pathservidor, "T"), "\"), Len(DBLet(RS!pathservidor, "T"))) & "\" & vAplic.Ejecutable
'            Version = a.GetFileVersion(Fichero)
'
'            SQL1 = "update aplicaciones set ultimaversion = " & DBSet(Version, "T")
'            SQL1 = SQL1 & " where idaplicacion = " & DBSet(RS!idaplicacion, "N")
'
'            conn.Execute SQL1
'
'            Set a = Nothing
'        End If
'
'        Set vAplic = Nothing
'        RS.MoveNext
'    Wend
'
'    Set RS = Nothing
'
'    ActualizarVersionesServidor = True
'    Exit Function
'
'eActualizarVersionesServidor:
'    If Err.Number <> 0 Then
'        Screen.MousePointer = vbDefault
'        MuestraError Err.Number, "Actualizando Versiones Servidor", Err.Description
'    End If
'End Function


Private Function VersionSuperior(v1 As String, v2 As String) As Boolean
Dim i As Integer
Dim J As Integer

    If InStr(1, v1, ".") = 0 Then
        If InStr(1, v2, ".") = 0 Then
            VersionSuperior = (CInt(Mid(v1, 1, Len(v1))) > CInt(Mid(v2, 1, Len(v2))))
        Else
            VersionSuperior = (Mid(v1, 1, Len(v1)) > Mid(v2, 1, InStr(1, v2, ".") - 1))
        End If
    Else
        If InStr(1, v2, ".") = 0 Then
            VersionSuperior = (Mid(v1, 1, InStr(1, v1, ".") - 1) > Mid(v2, 1, Len(v2)))
        Else
            If Mid(v1, 1, InStr(1, v1, ".") - 1) = Mid(v2, 1, InStr(1, v2, ".") - 1) Then
                VersionSuperior = VersionSuperior(Mid(v1, InStr(1, v1, ".") + 1, Len(v1)), Mid(v2, InStr(1, v2, ".") + 1, Len(v2)))
            Else
                VersionSuperior = (CInt(Mid(v1, 1, InStr(1, v1, ".") - 1)) > CInt(Mid(v2, 1, InStr(1, v2, ".") - 1)))
            End If
        End If
    End If
    
End Function



Private Function ComprobarVersionesPcPrevia(pc As String, ByRef Cad As String) As Boolean
Dim A As FileSystemObject

Dim Sql As String
Dim RS As ADODB.Recordset
Dim Aplic As String
Dim vAplic As CAplicacion
Dim Version As String
Dim fichero As String
Dim b As Boolean


    On Error GoTo eComprobarVersionesPCPrevia


    ComprobarVersionesPcPrevia = False

    Sql = "select pcscopia.* from pcscopia, pcs where ucase(pcs.nompc) = " & DBSet(UCase(pc), "T")
    Sql = Sql & " and pcscopia.idpcs = pcs.idpcs "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    
    Cad = ""
    
    While Not RS.EOF
        Set vAplic = New CAplicacion
        If vAplic.LeerDatos(RS!idaplicacion) Then
            
            
            Set A = New FileSystemObject
            
            fichero = DBLet(RS!pathcopia, "T") & "\" & vAplic.Ejecutable
            Version = A.GetFileVersion(fichero)
            
            If Version <> "" Then
                If VersionSuperior(vAplic.UltVers, Version) = True Then
                    Cad = Cad & DBLet(RS!idaplicacion, "N") & ":0|"
                End If
            Else
                Cad = Cad & DBLet(RS!idaplicacion, "N") & ":1|"
                b = False
            End If
            
            Set A = Nothing
        End If
        Set vAplic = Nothing
        RS.MoveNext
    Wend

    Set RS = Nothing

    ComprobarVersionesPcPrevia = True
    Exit Function

eComprobarVersionesPCPrevia:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
    
        MuestraError Err.Number, "Comprobando Versiones Previa", Err.Description
    End If
End Function

Public Function AplicacionesporActualizar(Cad As String) As String
Dim Cad1 As String
Dim Cad2 As String
Dim Resul As String
Dim J As Integer
Dim i As Integer
Dim Mens As String
Dim Mens1 As String
Dim longitud As Integer

Dim Aplic As String
Dim Situ As String

    AplicacionesporActualizar = ""

    Mens = "Se va a proceder a actualizar las siguientes aplicaciones: " & vbCrLf '& vbCrLf
    
    Mens1 = ""
    
    longitud = Len(Cad)
    
    i = 0
    Cad1 = Cad
    
    Aplicaciones = ""
    
    While Len(Cad1) <> 0
        i = InStr(1, Cad1, "|")
        
        If i <> 0 Then
            Cad2 = Mid(Cad1, 1, i)
            J = InStr(1, Cad2, ":")
            
            Aplic = Mid(Cad2, 1, J - 1)
            Situ = Mid(Cad2, J + 1, 1)
            
            If CInt(Situ) = 0 Then ' situacion sin actualizar
                Aplicaciones = Aplicaciones & Aplic & ","
                
                Mens = Mens & DevuelveDesdeBDNew(cPTours, "aplicaciones", "nombre", "idaplicacion", Aplic, "N")
                Mens = Mens & "    " '& vbCrLf
            Else  ' situacion de no poder actualizar
                Mens1 = Mens1 & DevuelveDesdeBDNew(cPTours, "aplicaciones", "nombre", "idaplicacion", Aplic, "N")
                Mens1 = Mens1 & "    " '& vbCrLf
            End If
            
            If Len(Cad1) <> i Then
                Cad1 = Mid(Cad1, i + 1, Len(Cad1))
            Else
                Cad1 = ""
            End If
        End If
    
    Wend

    Resul = ""
    If Aplicaciones <> "" Then
        Aplicaciones = Mid(Aplicaciones, 1, Len(Aplicaciones) - 1) ' quitamos la ultima coma
        Resul = Mens & vbCrLf '& vbCrLf
    End If
    If Mens1 <> "" Then
'        Resul = Resul & "Las siguientes aplicaciones no tienen versión y no se actualizarán: " & vbCrLf '& vbCrLf
        Resul = Resul & "No se encontró archivo ejecutable y no se actualizarán: " & vbCrLf '& vbCrLf
        Resul = Resul & Mens1
    End If
    
    HayPdte = True
    If Aplicaciones = "" Then HayPdte = False
        
    
    AplicacionesporActualizar = Resul
    
End Function

Private Function ActualizaCuentas(Tabla As String, IBAN As String, Banco As String, Sucursal As String, DC As String, CuentaBa As String) As Boolean
Dim A As FileSystemObject

Dim Sql As String
Dim RS As ADODB.Recordset
Dim Aplic As String
Dim vAplic As CAplicacion
Dim Version As String
Dim fichero As String
Dim b As Boolean
Dim IbanCorrecto As String
Dim DCCorrecto As String

Dim cta As String
Dim AA As String
Dim Sql2 As String
Dim NRegs As Integer
Dim Fic As String
Dim NF As Integer
Dim Atributos As Integer
Dim Cad1 As String

    On Error GoTo eActualizaCuentas

    Pb1.visible = False
    
    Me.Refresh
    DoEvents
    
    ActualizaCuentas = False

    conn.BeginTrans

    Sql = "select distinct " & IBAN & "," & Banco & "," & Sucursal & "," & DC & "," & CuentaBa & " from " & Tabla
    Sql = Sql & " where " & Banco & " >= 0 and " & Sucursal & " >= 0 and " & CuentaBa & " >0"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Sql2 = "select count(*) from (" & Sql & ") aaa "
    
    NRegs = TotalRegistros(Sql2)
    
    Pb1.Max = NRegs
    Pb1.visible = True
    Pb1.Value = 0
    
    b = True
    
    Cad1 = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    Fic = "c:\windows\temp\slog" & Mid(Cad1, 1, 4) & Mid(Cad1, 6, 2) & Mid(Cad1, 9, 2) & Mid(Cad1, 12, 2) & Mid(Cad1, 15, 2) & Mid(Cad1, 18, 2) & ".txt"
                
    NF = FreeFile
'    ' le quito el atributo de solo lectura
'    Atributos = GetAttr(Fic)
'    If (GetAttr(Fic) And vbReadOnly) <> 0 Then
'        SetAttr Fic, Atributos - vbReadOnly
'    End If
    
    
    
    Open Fic For Output As #NF
    
    Cad1 = "IBAN|Banco|Sucursal|DC|Cuenta|IBAN Calculado|DC Calculado|"
    Print #NF, Cad1
    
    While Not RS.EOF And b
        IncrementarProgresNew Pb1, 1
        DoEvents
    
        
        If DBLet(RS.Fields(1).Value, "N") = "" Or DBLet(RS.Fields(2).Value, "N") = "" Or DBLet(RS.Fields(4).Value, "N") = "" Then
        
        Else
            cta = Format(DBLet(RS.Fields(1).Value, "N"), "0000") & Format(DBLet(RS.Fields(2).Value, "N"), "0000") & "00" & Format(DBLet(RS.Fields(4).Value, "N"), "0000000000")
            If CDbl(cta) <> 0 Then
                DCCorrecto = DigitoControlCorrecto(cta)
                
                cta = Format(DBLet(RS.Fields(1).Value, "N"), "0000") & Format(DBLet(RS.Fields(2).Value, "N"), "0000") & Format(DCCorrecto, "00") & Format(DBLet(RS.Fields(4).Value, "N"), "0000000000")
            
                AA = ""
                If Len(DBLet(RS.Fields(0).Value, "N")) <> 0 Then AA = Mid(DBLet(RS.Fields(0).Value), 1, 2)
                    
                DevuelveIBAN2 AA, cta, cta
            
                IbanCorrecto = AA & cta
                
                
                If DCCorrecto <> DBLet(RS.Fields(3).Value, "T") Or IbanCorrecto <> DBLet(RS.Fields(0).Value, "T") Then
                
                    Sql2 = "update " & Tabla & " set "
                    Sql2 = Sql2 & Combo1(2).Text & " = " & DBSet(IbanCorrecto, "T") & ","
                    Sql2 = Sql2 & Combo1(3).Text & " = " & DBSet(RS.Fields(1).Value, "N") & ","
                    Sql2 = Sql2 & Combo1(4).Text & " = " & DBSet(RS.Fields(2).Value, "N") & ","
                    Sql2 = Sql2 & Combo1(5).Text & " = " & DBSet(Format(DCCorrecto, "00"), "T") & ","
                    Sql2 = Sql2 & Combo1(6).Text & " = " & DBSet(Format(RS.Fields(4).Value, "0000000000"), "T")
                    Sql2 = Sql2 & " where "
                    If IsNull(RS.Fields(0).Value) Then
                        Sql2 = Sql2 & Combo1(2).Text & " is null and "
                    Else
                        Sql2 = Sql2 & Combo1(2).Text & " = " & DBSet(RS.Fields(0).Value, "T") & " and "
                    End If
                    Sql2 = Sql2 & Combo1(3).Text & " = " & DBSet(RS.Fields(1).Value, "N") & " and "
                    Sql2 = Sql2 & Combo1(4).Text & " = " & DBSet(RS.Fields(2).Value, "N") & " and "
                    Sql2 = Sql2 & Combo1(5).Text & " = " & DBSet(RS.Fields(3).Value, "T") & " and "
                    Sql2 = Sql2 & Combo1(6).Text & " = " & DBSet(RS.Fields(4).Value, "T")
                    
                    conn.Execute Sql2
                    
                    Cad1 = DBLet(RS.Fields(0).Value, "N") & "|" & Format(DBLet(RS.Fields(1).Value, "N"), "0000") & "|" & Format(DBLet(RS.Fields(2).Value, "N"), "0000") & "|" & DBLet(RS.Fields(3).Value, "N") & "|" & DBLet(RS.Fields(4).Value, "N") & "|" & IbanCorrecto & "|" & DCCorrecto & "|"
                    
                    Print #NF, Cad1
                End If
                
            End If
        End If
        
        RS.MoveNext
    Wend
    Close #NF

    Set RS = Nothing
    
    conn.CommitTrans
    ActualizaCuentas = b
    Exit Function

eActualizaCuentas:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Actualizando Cuentas", Err.Description
    End If
End Function


