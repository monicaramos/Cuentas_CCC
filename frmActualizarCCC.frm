VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActualizarCCC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizador de CCC"
   ClientHeight    =   3555
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
   ScaleHeight     =   3555
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
      Top             =   2940
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   2310
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
      Top             =   2940
      Width           =   1275
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Actualizando tabla "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   2610
      Visible         =   0   'False
      Width           =   5865
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
Dim LTabla As Collection

Private Sub cmdAceptar_Click()
'Dim NomPc As String
Dim b As Boolean

    If Not DatosOk Then Exit Sub

    If Combo1(1).Text = "" Then Combo1_GotFocus (1)

    b = ActualizaCuentas(Combo1(1).Text)

    If b Then
        MsgBox "Proceso realizado correctamente", vbExclamation
    Else
        If LTabla.Count = 0 Then
            MsgBox "Esta Base de Datos no está contemplada. Revise."
        End If
    End If
    
    Me.Refresh
    DoEvents
 
    CmdSalir_Click
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Function DatosOk() As Boolean
    'DatosOk = Not (Combo1(0).Text = "" Or Combo1(1).Text = "" Or Combo1(2).Text = "" Or Combo1(3).Text = "" Or Combo1(4).Text = "" Or Combo1(5).Text = "" Or Combo1(6).Text = "")


    DatosOk = Not (Combo1(0).Text = "")

End Function


Private Sub Combo1_GotFocus(Index As Integer)
    If Index = 1 Then
        If Combo1(0).Text <> "" Then
        
            conn.Close
        
            AbrirConexion vConfig.User, vConfig.password, Combo1(0).Text
        
            CargaComboTabla
        End If
    
    End If
End Sub


Private Sub Combo1_LostFocus(Index As Integer)
    If Index = 0 Then Set LTabla = New Collection
    
    If Index = 1 Then
        If Combo1(Index).Text <> "" Then

            Set LTabla = New Collection
            LTabla.Add Combo1(1).Text
     
        End If
    End If
End Sub

Private Sub Form_Activate()
    If cerrar Then Unload Me
End Sub



Private Sub LimpiarCombosCampos()
Dim i As Integer

    For i = 1 To 1
        Combo1(i).Clear
    Next i
End Sub



Private Sub Form_Load()
Dim cad As String, Cad1 As String
Dim Mens As String
Dim b As Boolean

    AbrirConexion vConfig.User, vConfig.password, ""

    CargaComboBD

End Sub

Private Sub CargaComboBD()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer
Dim SQL As String
Dim RS As ADODB.Recordset
Dim BDatos As Collection
    
    
     Combo1(0).Clear
    
     Set RS = New ADODB.Recordset
        
     i = 0
     
     Set BDatos = New Collection
     
     RS.Open "SHOW DATABASES", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     While Not RS.EOF
     
        If CStr(RS.Fields(0)) <> "information_schema" And CStr(RS.Fields(0)) <> "test" And CStr(RS.Fields(0)) <> "mysql" And CStr(RS.Fields(0)) <> "usuarios" Then BDatos.Add CStr(RS.Fields(0))
     
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
Dim SQL As String
Dim RS As ADODB.Recordset

    
    
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
                    Else
                        If InStr(1, Combo1(0).Text, "aritaxi") <> 0 Then
                            LTabla.Add "sclien"
                            LTabla.Add "scliente"
                            LTabla.Add "sprove"
                            LTabla.Add "straba"
                            LTabla.Add "scafac"
                            LTabla.Add "scafaccli"
                            LTabla.Add "sbanpr"
                        End If
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
Dim SQL As String
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



Private Function ActualizaCuentas(Tabla As String) As Boolean
Dim A As FileSystemObject

Dim SQL As String
Dim RS As ADODB.Recordset
Dim Aplic As String
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
Dim Carpeta As String

Dim IBAN As String
Dim Banco As String
Dim Sucursal As String
Dim Cuentaba As String
Dim DC As String

Dim i As Integer

    On Error GoTo eActualizaCuentas

    Pb1.Visible = False
    
    Me.Refresh
    DoEvents
    
    ActualizaCuentas = False

    If LTabla.Count = 0 Then
        Exit Function
    End If


    conn.BeginTrans

    
    For i = 1 To LTabla.Count

        IBAN = "iban"
        Banco = "codbanco"
        Sucursal = "codsucur"
        DC = "digcontr"
        Cuentaba = "cuentaba"

        If InStr(1, Combo1(0).Text, "conta") <> 0 And (LTabla.Item(i) = "cuentas" Or LTabla.Item(i) = "spagop") Then
            IBAN = "iban"
            Banco = "entidad"
            Sucursal = "oficina"
            DC = "CC"
            Cuentaba = "cuentaba"
        End If

        SQL = "select distinct " & IBAN & "," & Banco & "," & Sucursal & "," & DC & "," & Cuentaba & " from " & LTabla.Item(i)
        SQL = SQL & " where " & Banco & " >= 0 and " & Sucursal & " >= 0 and " & Cuentaba & " >0"
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Sql2 = "select count(*) from (" & SQL & ") aaa "
        
        NRegs = TotalRegistros(Sql2)
        If NRegs <> 0 Then
            Pb1.Max = NRegs
            Pb1.Visible = True
            Pb1.Value = 0
            Label9(1).Visible = True
            Label9(1).Caption = "Actualizando Tabla " & LTabla.Item(i)
            
            b = True
            
            Cad1 = Format(Now, "yyyy-mm-dd hh:mm:ss")
            
            Carpeta = App.Path & "\logs"
            If Dir(Carpeta, vbDirectory) = "" Then
                MkDir Carpeta
            End If
            Fic = Carpeta & "\slog" & Mid(Cad1, 1, 4) & Mid(Cad1, 6, 2) & Mid(Cad1, 9, 2) & Mid(Cad1, 12, 2) & Mid(Cad1, 15, 2) & Mid(Cad1, 18, 2) & ".txt"
                        
            NF = FreeFile
            
            Open Fic For Output As #NF
            
            Cad1 = "BD|Tabla|IBAN|Banco|Sucursal|DC|Cuenta|IBAN Calculado|DC Calculado|"
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
                        
                            Sql2 = "update " & LTabla.Item(i) & " set "
                            Sql2 = Sql2 & IBAN & " = " & DBSet(IbanCorrecto, "T") & ","
                            Sql2 = Sql2 & Banco & " = " & DBSet(RS.Fields(1).Value, "N") & ","
                            Sql2 = Sql2 & Sucursal & " = " & DBSet(RS.Fields(2).Value, "N") & ","
                            Sql2 = Sql2 & DC & " = " & DBSet(Format(DCCorrecto, "00"), "T") & ","
                            Sql2 = Sql2 & Cuentaba & " = " & DBSet(Format(RS.Fields(4).Value, "0000000000"), "T")
                            Sql2 = Sql2 & " where "
                            If IsNull(RS.Fields(0).Value) Then
                                Sql2 = Sql2 & IBAN & " is null and "
                            Else
                                Sql2 = Sql2 & IBAN & " = " & DBSet(RS.Fields(0).Value, "T") & " and "
                            End If
                            Sql2 = Sql2 & Banco & " = " & DBSet(RS.Fields(1).Value, "N") & " and "
                            Sql2 = Sql2 & Sucursal & " = " & DBSet(RS.Fields(2).Value, "N") & " and "
    '                        Sql2 = Sql2 & DC & " = " & DBSet(RS.Fields(3).Value, "T","N") & " and "
                            Sql2 = Sql2 & Cuentaba & " = " & DBSet(RS.Fields(4).Value, "T")
                            
                            conn.Execute Sql2
                            
                            Cad1 = Combo1(0).Text & "|" & LTabla.Item(i) & "|" & DBLet(RS.Fields(0).Value, "N") & "|" & Format(DBLet(RS.Fields(1).Value, "N"), "0000") & "|" & Format(DBLet(RS.Fields(2).Value, "N"), "0000") & "|" & DBLet(RS.Fields(3).Value, "N") & "|" & DBLet(RS.Fields(4).Value, "N") & "|" & IbanCorrecto & "|" & DCCorrecto & "|"
                            
                            Print #NF, Cad1
                        End If
                        
                    End If
                End If
                
                RS.MoveNext
            Wend
            
            Cad1 = "Fin del proceso de la tabla " & Combo1(0).Text & "." & LTabla.Item(i)
            Print #NF, Cad1
    
            Close #NF
        
            Set RS = Nothing
    
            espera 2
        End If
    Next i
    
    conn.CommitTrans
    ActualizaCuentas = b
    
    Pb1.Visible = False
    Label9(1).Visible = False
    
    Exit Function

eActualizaCuentas:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Actualizando Cuentas", Err.Description
    End If
End Function


