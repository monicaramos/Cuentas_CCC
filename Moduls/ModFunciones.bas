Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit

Public Const ValorNulo = "Null"

Public Function CompForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox And Control.visible = True Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then Exit Function
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function CompForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm2 = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Carga = mTag.Cargar(Control)
                    If Carga = False Then
                        MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                        Exit Function
    
                    Else
                        If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                                MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                                Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    CompForm2 = True
End Function




'Public Function CampoSiguiente(ByRef formulario As Form, valor As Integer) As Control
'Dim Fin As Boolean
'Dim Control As Object
'
'On Error GoTo ECampoSiguiente
'
'    'Debug.Print "Llamada:  " & Valor
'    'Vemos cual es el siguiente
'    Do
'        valor = valor + 1
'        For Each Control In formulario.Controls
'            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
'            'Si es texto monta esta parte de sql
'            If Control.TabIndex = valor Then
'                    Set CampoSiguiente = Control
'                    Fin = True
'                    Exit For
'            End If
'        Next Control
'        If Not Fin Then
'            valor = -1
'        End If
'    Loop Until Fin
'    Exit Function
'ECampoSiguiente:
'    Set CampoSiguiente = Nothing
'    Err.Clear
'End Function



'-----------------------------------
Public Function ValorParaSQL(Valor, ByRef vtag As CTag) As String
Dim Dev As String
Dim d As Single
Dim i As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vtag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Or InStr(1, Valor, ".") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004

                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else

                    V = CSng(Valor)
                    Valor = V
                End If
            Else

            End If
            Dev = TransformaComasPuntos(CStr(Valor))

        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
            
        Case "H"
            Dev = "'" & Format(Valor, "hh:mm:ss") & "'"
        
        Case "FHH"
            Dev = DBSet(Valor, "FH")
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select

    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vtag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function


Public Function InsertarDesdeForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                    Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.Tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    
    conn.Execute cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function InsertarDesdeForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm2 = False
    Der = ""
    Izda = ""
    
    For Each Control In formulario.Controls
    
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.Value = 1 Then
                        cad = "1"
                        Else
                        cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                    Der = Der & cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & ","
                            Izda = Izda & "" & mTag.columna & ""
                            cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
'                        If Control.Value Then
'                            If Izda <> "" Then Izda = Izda & ","
'                            Izda = Izda & "" & mTag.columna & ""
'                            cad = Control.index
'                            If Der <> "" Then Der = Der & ","
'                            Der = Der & cad
'                        End If
                        If Izda <> "" Then Izda = Izda & ","
                        Izda = Izda & "" & mTag.columna & ""
                        
                        'Parte VALUES
                        If Control.visible Then
                            cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            cad = ValorNulo
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.Tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute cad, , adCmdText
    
     ' ### [Monica] 18/12/2006
    CadenaCambio = cad
   
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function CadenaInsertarDesdeForm(ByRef formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la función.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                    Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.Tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim i As Integer

    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) And (Control.visible = True) Then
'                If TypeOf control Is TextBox Then

            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.columna <> "" Then
                        campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) And (Control.visible = True) Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                If IsNull(Valor) Then Valor = 0
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.columna
                    Valor = DBLet(vData.Recordset.Fields(campo))
                    i = 0
                    For i = 0 To Control.ListCount - 1
                        If Control.ItemData(i) = Val(Valor) Then
                            Control.ListIndex = i
                            Exit For
                        End If
                    Next i
                    If i = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function PonerCamposForma2(ByRef formulario As Form, ByRef vData As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim i As Integer

    Set mTag = New CTag
    PonerCamposForma2 = False

    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    If mTag.Cargado Then
                        'Columna en la BD
                        If mTag.columna <> "" Then
                            campo = mTag.columna
                            If mTag.Vacio = "S" Then
                                Valor = DBLet(vData.Recordset.Fields(campo))
                            Else
                                Valor = vData.Recordset.Fields(campo)
                            End If
                            If mTag.Formato <> "" And CStr(Valor) <> "" Then
                                If mTag.TipoDato = "N" Then
                                    'Es numerico, entonces formatearemos y sustituiremos
                                    ' La coma por el punto
                                    cad = Format(Valor, mTag.Formato)
                                    'Antiguo
                                    'Control.Text = TransformaComasPuntos(cad)
                                    'nuevo
                                    Control.Text = cad
                                Else
                                    Control.Text = Format(Valor, mTag.Formato)
                                End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                    End If
                    If IsNull(Valor) Then Valor = 0
                    Control.Value = Valor
                End If
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        campo = mTag.columna
                        Valor = DBLet(vData.Recordset.Fields(campo))
                        i = 0
                        For i = 0 To Control.ListCount - 1
                            If Control.ItemData(i) = Val(Valor) Then
                                Control.ListIndex = i
                                Exit For
                            End If
                        Next i
                        If i = Control.ListCount Then Control.ListIndex = -1
                    End If 'de cargado
                End If
            End If 'de <>""
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                        If IsNull(Valor) Then Valor = 0
                        If Control.Index = Valor Then
                            Control.Value = True
                        Else
                            Control.Value = False
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                        If IsNull(Valor) Then Valor = Now
                        Control.Value = Format(Valor, mTag.Formato)
                    End If
                End If
            End If
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma2 = True
Exit Function
EPonerCamposForma2:
    MuestraError Err.Number, "Poner campos formulario 2. "
End Function


Public Function ForaGrid(ByRef formulari As Form, ByRef vGrid As DataGrid, Control As Object) As Boolean
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim camp As String  'Camp en la BDA
Dim i As Integer

    Set mTag = New CTag
    ForaGrid = False

    If (TypeOf Control Is TextBox) Then 'text
        mTag.Cargar Control
        If Control.Tag <> "" Then
            If mTag.Cargado Then
                If mTag.columna <> "" Then
                    camp = mTag.columna
                    If mTag.Vacio = "S" Then
                        Valor = DBLet(vGrid.Columns(camp).Text)
                        'valor = DBLet(vGrid.Recordset.Fields(campo))
                    Else
                        'valor = vGrid.Columns!camp
                        Valor = vGrid.Columns(camp).Text
                    End If
                    If mTag.Formato <> "" And CStr(Valor) <> "" Then
                        If mTag.TipoDato = "N" Then
                            cad = Format(Valor, mTag.Formato)
                            Control.Text = cad
                        Else
                            Control.Text = Format(Valor, mTag.Formato)
                        End If
                    Else
                        Control.Text = Valor
                    End If
                End If
            End If
        End If

'        'CheckBOX
'        ElseIf (TypeOf Control Is CheckBox) Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        'Columna en la BD
'                        campo = mTag.columna
'                        valor = vData.Recordset.Fields(campo)
'                        Else
'                            valor = 0
'                    End If
'                    If IsNull(valor) Then valor = 0
'                    Control.Value = valor
'                End If
'            End If
'
'         'COMBOBOX
'         ElseIf (TypeOf Control Is ComboBox) Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        campo = mTag.columna
'                        valor = DBLet(vData.Recordset.Fields(campo))
'                        i = 0
'                        For i = 0 To Control.ListCount - 1
'                            If Control.ItemData(i) = Val(valor) Then
'                                Control.ListIndex = i
'                                Exit For
'                            End If
'                        Next i
'                        If i = Control.ListCount Then Control.ListIndex = -1
'                    End If 'de cargado
'                End If
'            End If 'de <>""
    End If

    'Veremos que tal
    ForaGrid = True
Exit Function
EPosarCampsGrid:
    MuestraError Err.Number, "Poner campos grid. "
End Function


'Public Function PonerCamposFormaFrame(ByRef formulario As Form, NomTxtBox As String, ByRef vData As Adodc, Optional NomCheck As String, Optional NomCombo As String) As Boolean
'Dim Control As Object
'Dim mTag As CTag
'Dim cad As String
'Dim valor As Variant
'Dim campo As String  'Campo en la base de datos
'Dim i As Integer
'
'    Set mTag = New CTag
'    PonerCamposFormaFrame = False
'
'
'        For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox And Control.Visible = True And Control.Name = NomTxtBox Then
'            'Comprobamos que tenga tag
'            mTag.Cargar Control
''            Debug.Print Control.Parent
'            If Control.Tag <> "" Then
'                If mTag.Cargado Then
'                    'Columna en la BD
'                    If mTag.Columna <> "" Then
'                        campo = mTag.Columna
'                        If mTag.Vacio = "S" Then
'                            valor = DBLet(vData.Recordset.Fields(campo))
'                        Else
'                            valor = vData.Recordset.Fields(campo)
'                        End If
'                        If mTag.Formato <> "" And CStr(valor) <> "" Then
'                            If mTag.TipoDato = "N" Then
'                                'Es numerico, entonces formatearemos y sustituiremos
'                                ' La coma por el punto
'                                cad = Format(valor, mTag.Formato)
'                                'Antiguo
'                                'Control.Text = TransformaComasPuntos(cad)
'                                'nuevo
'                                Control.Text = cad
'                            Else
'                                Control.Text = Format(valor, mTag.Formato)
'                            End If
'                        Else
'                            Control.Text = valor
'                        End If
'                    End If
'                End If
'            End If
'        'CheckBOX
'        ElseIf TypeOf Control Is CheckBox And Control.Visible = True And Control.Name = NomCheck Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    'Columna en la BD
'                    campo = mTag.Columna
'                    valor = vData.Recordset.Fields(campo)
'                    Else
'                        valor = 0
'                End If
'                Control.Value = valor
'            End If
'
'         'COMBOBOX
'         ElseIf TypeOf Control Is ComboBox And Control.Visible = True And Control.Name = NomCombo Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    campo = mTag.Columna
'                    valor = vData.Recordset.Fields(campo)
'                    i = 0
'                    For i = 0 To Control.ListCount - 1
'                        If Control.ItemData(i) = Val(valor) Then
'                            Control.ListIndex = i
'                            Exit For
'                        End If
'                    Next i
'                    If i = Control.ListCount Then Control.ListIndex = -1
'                End If 'de cargado
'            End If 'de <>""
'        End If
'
'    Next Control
'
'    'Veremos que tal
'    PonerCamposFormaFrame = True
'Exit Function
'EPonerCamposForma:
'    MuestraError Err.Number, "Poner campos formulario. "
'End Function


Private Function ObtenerMaximoMinimo(vSQL As String, Optional vBD As Byte) As String
Dim RS As Recordset
    ObtenerMaximoMinimo = ""
    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.EOF) Then
            ObtenerMaximoMinimo = CStr(RS.Fields(0))
        End If
    End If
    RS.Close
    Set RS = Nothing
End Function


'====DAVID
'Public Function ObtenerBusqueda(ByRef formulario As Form) As String
'    Dim Control As Object
'    Dim Carga As Boolean
'    Dim mTag As CTag
'    Dim Aux As String
'    Dim cad As String
'    Dim SQL As String
'    Dim tabla As String
'    Dim RC As Byte
'
'    On Error GoTo EObtenerBusqueda
'
'    'Exit Function
'    Set mTag = New CTag
'    ObtenerBusqueda = ""
'    SQL = ""
'
'    'Recorremos los text en busca de ">>" o "<<"
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If Aux = ">>" Or Aux = "<<" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'                    If Aux = ">>" Then
'                        cad = " MAX(" & mTag.Columna & ")"
'                    Else
'                        cad = " MIN(" & mTag.Columna & ")"
'                    End If
'                    SQL = "Select " & cad & " from " & mTag.tabla
'                    SQL = ObtenerMaximoMinimo(SQL)
'                    Select Case mTag.TipoDato
'                    Case "N"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
'                    Case "F"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
'                    Case Else
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
'                    End Select
'                    SQL = "(" & SQL & ")"
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los text en busca del NULL
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If UCase(Aux) = "NULL" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'
'                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
'                    SQL = "(" & SQL & ")"
'                    Control.Text = ""
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los textbox
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            'Cargamos el tag
'            Carga = mTag.Cargar(Control)
'            If Carga Then
'                If mTag.Cargado Then
'                    Aux = Trim(Control.Text)
'                    If Aux <> "" Then
'                        If mTag.tabla <> "" Then
'                            tabla = mTag.tabla & "."
'                        Else
'                            tabla = ""
'                        End If
'                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad)
'                    If RC = 0 Then
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'            Else
'                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
'                Exit Function
'            End If
'
'        'COMBO BOX
'        ElseIf TypeOf Control Is ComboBox Then
'            mTag.Cargar Control
'            If mTag.Cargado Then
'                If Control.ListIndex > -1 Then
'                    If mTag.TipoDato <> "T" Then
'                        cad = Control.ItemData(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    Else
'                        cad = Control.List(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = '" & cad & "'"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'
'
'        'CHECK
'        ElseIf TypeOf Control Is CheckBox Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    If Control.Value = 1 Then
'                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'        End If
'
'
'    Next Control
'    ObtenerBusqueda = SQL
'Exit Function
'EObtenerBusqueda:
'    ObtenerBusqueda = ""
'    MuestraError Err.Number, "Obtener búsqueda. "
'End Function

Public Function ObtenerBusqueda(ByRef formulario As Form, Optional CHECK As String, Optional vBD As Byte, Optional cadwhere As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim Tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                If Control.Tag <> "" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Aux = ">>" Then
                            cad = " MAX("
                        Else
                            cad = " MIN("
                        End If
                        'monica
                        Select Case mTag.TipoDato
                            Case "FHF"
                                cad = cad & "date(" & mTag.columna & "))"
                            Case "FHH"
                                cad = cad & "time(" & mTag.columna & "))"
                            Case Else
                                cad = cad & mTag.columna & ")"
                        End Select
                        
                        Sql = "Select " & cad & " from " & mTag.Tabla
                        If cadwhere <> "" Then Sql = Sql & " WHERE " & cadwhere
                        Sql = ObtenerMaximoMinimo(Sql, vBD)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHF"
                            Sql = "date(" & mTag.Tabla & "." & mTag.columna & ") = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHH"
                            Sql = "time(" & mTag.Tabla & "." & mTag.columna & ") = '" & Format(Sql, "hh:mm:ss") & "'"
                        Case Else
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Sql & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

    

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                'Cargamos el tag
                Carga = mTag.Cargar(Control)
                If Carga Then
'                    Debug.Print Control.Tag
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, cad)
                        If RC = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex > -1 Then
                        If mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 15/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        cad = Control.Value
                        cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional CHECK As String, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim Tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            cad = " MAX(" & mTag.columna & ")"
                        Else
                            cad = " MIN(" & mTag.columna & ")"
                        End If
                        Sql = "Select " & cad & " from " & mTag.Tabla
                        Sql = ObtenerMaximoMinimo(Sql)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case Else
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Sql & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
          If Control.Tag <> "" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, cad)
                        If RC = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de Cèsar, no te sentit passar-li un control que no té TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            cad = Control.ItemData(Control.ListIndex)
                            cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    ' añadido 12022007
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        cad = Control.Value
                        cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function


Public Function ModificaDesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                             cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"

                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                         cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                    Else
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
'
'
'                   If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                   'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                   cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Function ModificaDesdeFormulario2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                 cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Control.Value = 1 Then
                        Aux = "TRUE"
                    Else
                        Aux = "FALSE"
                    End If
                    If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'Esta es para access
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.ListIndex = -1 Then
                            Aux = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Aux = Control.ItemData(Control.ListIndex)
                        Else
                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                             cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
'
'
'                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            Aux = Control.Index
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                              If mTag.EsClave Then
                                  'Lo pondremos para el WHERE
                                   If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                   cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                              Else
                                  If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                  cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                              End If
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
'                        If Control.Value Then
                         If mTag.columna <> "" Then
'                            Aux = Control.index
                            If Control.visible Then
                                Aux = ValorParaSQL(Control.Value, mTag)
                            Else
                                Aux = ValorNulo
                            End If
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    conn.Execute Aux, , adCmdText

    ' ### [Monica] 18/12/2006
    CadenaCambio = cadUPDATE

    ModificaDesdeFormulario2 = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar 2. " & Err.Description
End Function


Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                cad = TransformaPuntosComas(vTex.Text)
                cad = Format(cad, mTag.Formato)
                vTex.Text = cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


Public Function FormatoCampo(ByRef vTex As TextBox) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String
    
    On Error GoTo EFormatoCampo

    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        FormatoCampo = mTag.Formato
    End If
    
EFormatoCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'Añade: CESAR
'Para utilizalo en el arreglaGrid
Public Function FormatoCampo2(ByRef objec As Object) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String

    On Error GoTo EFormatoCampo2

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        FormatoCampo2 = mTag.Formato
    End If
    
EFormatoCampo2:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef cadena As String, Orden As Integer, Optional separador As String) As String
Dim i As Integer
Dim J As Integer
Dim Cont As Integer
Dim cad As String

    i = 0
    Cont = 1
    cad = ""
    Do
        J = i + 1
        If separador = "" Then
            i = InStr(J, cadena, "|")
        Else
            i = InStr(J, cadena, separador)
        End If
        If i > 0 Then
            If Cont = Orden Then
                cad = Mid(cadena, J, i - J)
                i = Len(cadena) 'Para salir del bucle
                Else
                    Cont = Cont + 1
            End If
        End If
    Loop Until i = 0
    RecuperaValor = cad
End Function



'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim i As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneral
'bol = vSesion.Nivel < 2

'Añadir, modificar y borrar deshabilitados si no nivel
With formulario
    For i = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(i).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(i).Tag)
            If J < vSesion.Nivel Then
                .Toolbar1.Buttons(i).Enabled = False
            End If
        End If
    Next i
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub


Public Sub PonerModoMenuGral(ByRef formulario As Form, activo As Boolean)
Dim i As Integer
'Dim j As Integer

On Error GoTo PonerModoMenuGral

'Añadir, modificar y borrar deshabilitados si no Modo
    With formulario
        For i = 1 To .Toolbar1.Buttons.Count
            Select Case .Toolbar1.Buttons(i).ToolTipText
                Case "Nuevo"
                    .Toolbar1.Buttons(i).visible = Not .DeConsulta
                Case "Modificar", "Eliminar", "Imprimir"
                    .Toolbar1.Buttons(i).visible = Not .DeConsulta
                    .Toolbar1.Buttons(i).Enabled = activo
'                Case "Modificar"
'                Case "Eliminar"
'                Case "Imprimir"
            End Select
        Next i
        
        
        'El menu Visible
        .mnModificar.visible = Not .DeConsulta
        .mnEliminar.visible = Not .DeConsulta
        'El menu activo
        .mnModificar.Enabled = activo
        .mnEliminar.Enabled = activo
    End With
    
    
Exit Sub
PonerModoMenuGral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub

Public Sub PonerOpcionesMenuGeneralNew(formulario As Form)
Dim Control As Object
Dim i As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneralNew
'bol = vSesion.Nivel < 2
'Añadir, modificar y borrar deshabilitados si no nivel
    For Each Control In formulario.Controls
'        Debug.Print Control.Name
        
        If Mid(Control.Name, 1, 2) = "mn" And Mid(Control.Name, 1, 7) <> "mnBarra" _
           And Control.Name <> "mnOpciones" Then
            J = Val(Control.HelpContextID)
            If J < vSesion.Nivel And J <> 0 Then
                Control.Enabled = False
            End If
        End If
    Next Control

Exit Sub
EPonerOpcionesMenuGeneralNew:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String
Dim i As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadwhere = Claves
    'Construimos el SQL
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    conn.Execute Aux, , adCmdText

ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function BLOQUEADesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then

                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                         cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control

    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadwhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function BLOQUEADesdeFormulario2(ByRef formulario As Form, ByRef ado As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim AntiguoCursor As Byte
Dim nomcamp As String

    On Error GoTo EBLOQUEADesdeFormulario2
    
    BLOQUEADesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If (TypeOf Control Is TextBox) Or (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        'Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            Aux = ValorParaSQL(CStr(ado.Recordset.Fields(mTag.columna)), mTag)
                            'Lo pondremos para el WHERE
                             If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                             cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadwhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario2 = True
    End If
    
EBLOQUEADesdeFormulario2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla 2"
'        BLOQUEADesdeFormulario2 = False
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistro(cadTABLA As String, cadwhere As String) As Boolean
Dim Aux As String

    On Error GoTo EBloqueaRegistro
        
    BloqueaRegistro = False
    Aux = "select * FROM " & cadTABLA
    Aux = Aux & " WHERE " & cadwhere & " FOR UPDATE"

    'Intenteamos bloquear
    PreparaBloquear
    conn.Execute Aux, , adCmdText
    BloqueaRegistro = True

EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control

    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
'        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & AuxDef & """)"
        conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
'        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.tabla & "'"
        conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function


'====================== LAURA

Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function



Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vtag As CTag
Dim devuelve As String

    On Error GoTo ErrExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vtag = New CTag
            If vtag.Cargar(T) Then
'                If vtag.EsClave Then
                    devuelve = DevuelveDesdeBDNew(cPTours, vtag.Tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
                    If devuelve <> "" Then
    '                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                        MsgBox "Ya existe el " & vtag.Nombre & ": " & T.Text, vbExclamation
                        ExisteCP = True
                        PonerFoco T
                    End If
'                End If
            End If
            Set vtag = Nothing
        End If
    End If
    Exit Function
    
ErrExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar código.", Err.Description
End Function




Public Function TotalRegistros(vSQL As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then TotalRegistros = RS.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function


' del ariges de Laura

Public Function BloqueoManual(cadTABLA As String, cadwhere As String)
Dim Aux As String

On Error GoTo EBLOQ

    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vSesion.Codigo & ",'" & cadTABLA
        Aux = Aux & "',""" & cadwhere & """)"
        conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
'        Else
'            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTABLA As String) As Boolean
Dim Sql As String
'Solo me interesa la tabla
On Error Resume Next

        Sql = "DELETE FROM zbloqueos WHERE codusu=" & vSesion.Codigo & " and tabla='" & cadTABLA & "'"
        conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function

Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim ent As Integer
Dim cad As String

  ' Comprobaciones

  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
    Exit Function
  End If

  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If

  ' Redondeo.

  cad = "0"
  If NumDigitsAfterDecimals <> 0 Then cad = cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Val(TransformaComasPuntos(Format(Number, cad)))

End Function

Public Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function


Public Function BorrarTablaIntermedia() As Boolean
Dim Sql As String
    
    On Error GoTo eBorrarTablaIntermedia
    
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    conn.Execute Sql
    
    BorrarTablaIntermedia = True
    Exit Function

eBorrarTablaIntermedia:
    BorrarTablaIntermedia = False

End Function

Public Function Repeat(cad As String, nveces As Long) As String
Dim Cad2 As String
Dim i As Long

    Cad2 = ""
    For i = 1 To nveces
        Cad2 = Cad2 & cad
    Next i
    Repeat = Cad2

End Function
