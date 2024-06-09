VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGenDAO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador DAO"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtArchivoDAO 
      Height          =   285
      Left            =   240
      TabIndex        =   37
      Top             =   2160
      Width           =   9375
   End
   Begin VB.CommandButton cmdArchivoDAO 
      Height          =   255
      Left            =   9360
      Picture         =   "frmGenDAO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton cmdGenerarDAO 
      Caption         =   "Generar"
      Height          =   375
      Left            =   9840
      TabIndex        =   35
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtUrl 
      Height          =   285
      Left            =   9840
      TabIndex        =   33
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CheckBox chkNullable 
      Caption         =   "Nullable"
      Height          =   255
      Left            =   7920
      TabIndex        =   32
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox chkAutoIncrement 
      Caption         =   "AutoIncrement"
      Height          =   255
      Left            =   7920
      TabIndex        =   31
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdDesconectar 
      Caption         =   "Desconectar"
      Height          =   375
      Left            =   4080
      TabIndex        =   30
      Top             =   960
      Width           =   1695
   End
   Begin VB.CheckBox chkFind 
      Caption         =   "Find"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ComboBox cboTablas 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   5535
   End
   Begin VB.ComboBox cboDriver 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdConectar 
      Caption         =   "Conectar"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtEndpoint 
      Height          =   285
      Left            =   9840
      TabIndex        =   15
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   9840
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   9840
      TabIndex        =   16
      Top             =   1560
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdCampos 
      Height          =   4335
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7646
      _Version        =   393216
   End
   Begin VB.CheckBox chkPrimaria 
      Caption         =   "Clave Primaria"
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdlArchivo 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdArchivo 
      Height          =   255
      Left            =   9360
      Picture         =   "frmGenDAO.frx":0071
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   9375
   End
   Begin VB.TextBox txtAtributo 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   17
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Archivo REPOSITORY"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   38
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Url"
      Height          =   195
      Index           =   5
      Left            =   9840
      TabIndex        =   34
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Tablas"
      Height          =   195
      Index           =   12
      Left            =   6000
      TabIndex        =   29
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Driver"
      Height          =   195
      Index           =   11
      Left            =   6000
      TabIndex        =   28
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   27
      Top             =   720
      Width           =   690
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   9
      Left            =   4080
      TabIndex        =   26
      Top             =   120
      Width           =   690
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "User"
      Height          =   195
      Index           =   8
      Left            =   2160
      TabIndex        =   25
      Top             =   120
      Width           =   330
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "IP"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   150
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Endpoint"
      Height          =   195
      Index           =   6
      Left            =   9840
      TabIndex        =   23
      Top             =   4560
      Width           =   630
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Atributos"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   21
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Archivo MODEL"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Atributo"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   2520
      Width           =   540
   End
End
Attribute VB_Name = "frmGenDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tipo As New clsTipo
Private atributo As New clsAtributo
Private dao As New clsDAO

Private DB As New ADODB.Connection

Private Sub cboTablas_Click()
Dim rstQuery As ADODB.Recordset

On Error Resume Next

    If Me.cboTablas.ListIndex < 0 Then Exit Sub
    
    Me.txtArchivo.Text = ""
    Me.txtArchivoDAO.Text = ""
    Me.txtEndpoint.Text = Me.cboTablas.Text
    
    dao.clean
    
    Set rstQuery = DB.OpenSchema(adSchemaColumns, Array(Me.txtDatabase.Text, Empty, Me.cboTablas.Text))
    
    Do While Not rstQuery.EOF
        Set atributo = New clsAtributo
        
        atributo.atributo = firstUpper(Replace(modConv.field2Attribute(LCase(rstQuery.Fields("COLUMN_NAME"))), "id", "Id"))
        atributo.tipo = modConv.typeADO2Visual(rstQuery.Fields("DATA_TYPE"))
        atributo.primaria = False
        atributo.find = False
        atributo.autoIncrement = rstQuery.Fields("COLUMN_NAME").Properties("ISAUTOINCREMENT")
        atributo.nullable = rstQuery.Fields("IS_NULLABLE")
        
        dao.atributos.Add atributo, atributo.atributo
    
        rstQuery.MoveNext
    Loop
    
    rstQuery.Close
    
    Set rstQuery = DB.OpenSchema(adSchemaPrimaryKeys, Array(Me.txtDatabase.Text, Empty, Me.cboTablas.Text))

    Do While Not rstQuery.EOF
        Set atributo = dao.atributos.Item(rstQuery.Fields("COLUMN_NAME"))
        atributo.primaria = True

        rstQuery.MoveNext
    Loop

    rstQuery.Close
    
    fillGrid
    
End Sub

Private Sub cmdAgregar_Click()

    If Me.cboTipo.ListIndex < 0 Then Exit Sub
    
    If Not modCollection.collectionExistElement(dao.atributos, Me.txtAtributo.Text) Then Set atributo = New clsAtributo
    
    atributo.atributo = firstUpper(Me.txtAtributo.Text)
    atributo.tipo = Me.cboTipo.Text
    atributo.primaria = IIf(Me.chkPrimaria.Value = 1, True, False)
    atributo.find = IIf(Me.chkFind.Value = 1, True, False)
    atributo.nullable = IIf(Me.chkNullable.Value = 1, True, False)
    atributo.autoIncrement = IIf(Me.chkAutoIncrement.Value = 1, True, False)
    
    If Trim(Me.txtAtributo.Text) = "" Then
        dao.atributos.Add atributo
    Else
        If Not modCollection.collectionExistElement(dao.atributos, atributo.atributo) Then dao.atributos.Add atributo, atributo.atributo
    End If
    
    fillGrid
    
    Me.txtAtributo.SetFocus
    
End Sub

Private Sub cmdArchivo_Click()

    Me.cdlArchivo.Filter = "Class (*.cls)|*.cls"
    
    Me.cdlArchivo.ShowSave
    
    Me.txtArchivo.Text = Me.cdlArchivo.FileName

End Sub

Private Sub cmdArchivoDAO_Click()

    Me.cdlArchivo.Filter = "Class (*.cls)|*.cls"
    
    Me.cdlArchivo.ShowSave
    
    Me.txtArchivoDAO.Text = Me.cdlArchivo.FileName

End Sub

Private Sub cmdConectar_Click()
Dim rstQuery As ADODB.Recordset

    If DB.State = adStateOpen Then Exit Sub
    DB.ConnectionString = "driver={" & Me.cboDriver.Text & "};server=" & Me.txtIP.Text & ";database=" & Me.txtDatabase.Text & ";user=" & Me.txtUser.Text & ";password=" & Me.txtPwd.Text & ";Trusted_Connection=yes;"
    DB.CursorLocation = adUseServer
    DB.Open
    Set rstQuery = DB.OpenSchema(adSchemaTables)
    
    Do While Not rstQuery.EOF
        Me.cboTablas.AddItem rstQuery.Fields("TABLE_NAME")
        rstQuery.MoveNext
    Loop
    
    rstQuery.Close
    
    If Me.cboTablas.ListCount > 0 Then Me.cboTablas.ListIndex = 0
    
End Sub

Private Sub cmdDesconectar_Click()

    DB.Close
    
    Me.cboTablas.Clear
    
End Sub

Private Sub cmdEliminar_Click()

    If Me.grdCampos.Rows = 1 Then Exit Sub
    
    dao.atributos.Remove Me.grdCampos.Row
    
    fillGrid
    
End Sub

Private Sub cmdGenerar_Click()
Dim strClase As String

Dim strLinea As Variant

Dim intCiclo As Integer

Dim blnNombre As Boolean

    If Me.txtArchivo.Text = "" Then
        MsgBox "ERROR: Falta Archivo"
        Exit Sub
    End If
    
    If Me.grdCampos.Rows = 1 Then
        MsgBox "ERROR: Sin ATRIBUTOS"
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    intCiclo = Len(Me.txtArchivo.Text) - 4
    
    blnNombre = False
    Do While intCiclo > 0 And Not blnNombre
        If Mid(Me.txtArchivo.Text, intCiclo, 1) = "\" Then
            blnNombre = True
        Else
            strClase = Mid(Me.txtArchivo.Text, intCiclo, 1) & strClase
        End If
        intCiclo = intCiclo - 1
    Loop
    
    dao.file = modConv.parseFilename(Me.txtArchivo.Text, dao.path)
    dao.className = strClase
    dao.endpoint = Me.txtEndpoint.Text
    dao.writeHeader
    dao.writeAttributes
    dao.writeInitialize
    dao.writeGetterSetter
    dao.writeClone
    dao.writeMakeParams
    dao.writeFillObject
    dao.closeFile
    
    Me.MousePointer = 0
    
    MsgBox "Generación TERMINADA"
    
End Sub

Private Sub cmdGenerarDAO_Click()
Dim strClase As String

Dim strLinea As Variant

Dim intCiclo As Integer

Dim blnNombre As Boolean

    If Me.txtArchivoDAO.Text = "" Then
        MsgBox "ERROR: Falta Archivo"
        Exit Sub
    End If
    
    If Me.grdCampos.Rows = 1 Then
        MsgBox "ERROR: Sin ATRIBUTOS"
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    intCiclo = Len(Me.txtArchivoDAO.Text) - 4
    
    blnNombre = False
    Do While intCiclo > 0 And Not blnNombre
        If Mid(Me.txtArchivoDAO.Text, intCiclo, 1) = "\" Then
            blnNombre = True
        Else
            strClase = Mid(Me.txtArchivoDAO.Text, intCiclo, 1) & strClase
        End If
        intCiclo = intCiclo - 1
    Loop
    
    dao.file = modConv.parseFilename(Me.txtArchivoDAO.Text, dao.path)
    dao.className = strClase
    dao.endpoint = Me.txtEndpoint.Text
    dao.writeHeader
    dao.url = Me.txtUrl.Text
    
    intCiclo = Len(Me.txtArchivo.Text) - 4
    strClase = ""
    
    blnNombre = False
    Do While intCiclo > 0 And Not blnNombre
        If Mid(Me.txtArchivo.Text, intCiclo, 1) = "\" Then
            blnNombre = True
        Else
            strClase = Mid(Me.txtArchivo.Text, intCiclo, 1) & strClase
        End If
        intCiclo = intCiclo - 1
    Loop
    
    dao.className = strClase
    dao.writeFindREST
    dao.writeFindByPrimaryKey
    dao.writeDelete
    dao.writeSave
    dao.writeExist
    dao.writeAdd
    dao.writeUpdate
    dao.writeCollectionAll
    dao.writeFillCombo
    dao.writeFillList
    dao.closeFile
    
    Me.MousePointer = 0
    
    MsgBox "Generación TERMINADA"
    
End Sub

Private Sub cmdSalir_Click()

    End
    
End Sub

Private Sub Form_Load()
Dim varTitulos As Variant
Dim varAnchos As Variant

    tipo.fillCombo Me.cboTipo
    
    varTitulos = Array("Atributo", "Tipo", "Primary", "Find", "Nullable", "AutoIncrement")
    varAnchos = Array(2000, 2000, 1000, 1000, 1000, 1000)
    
    makeGrid Me.grdCampos, varTitulos, varAnchos, 0, 1, flexSelectionByRow
    
    Me.cboDriver.AddItem "MySQL ODBC 5.1 Driver"
    Me.cboDriver.AddItem "SQL Native Client"
    Me.cboDriver.ListIndex = 0
    Me.txtUrl.Text = "url"
    
End Sub

Private Sub fillGrid()
Dim strLinea As String

    Me.grdCampos.Rows = 1
    
    For Each atributo In dao.atributos
        strLinea = atributo.atributo & Chr(9)
        strLinea = strLinea & atributo.tipo & Chr(9)
        strLinea = strLinea & IIf(atributo.primaria, "Si", "No") & Chr(9)
        strLinea = strLinea & IIf(atributo.find, "Si", "No") & Chr(9)
        strLinea = strLinea & IIf(atributo.nullable, "Si", "No") & Chr(9)
        strLinea = strLinea & IIf(atributo.autoIncrement, "Si", "No") & Chr(9)
        
        Me.grdCampos.AddItem strLinea
    Next
    
End Sub

Private Sub grdCampos_Click()

On Error Resume Next

    If Me.grdCampos.Row < 1 Then Exit Sub
    
    Set atributo = dao.atributos.Item(Me.grdCampos.Row)
    
    With atributo
        Me.txtAtributo.Text = .atributo
        Me.cboTipo.Text = .tipo
        Me.chkPrimaria.Value = IIf(.primaria, 1, 0)
        Me.chkFind.Value = IIf(.find, 1, 0)
        Me.chkNullable.Value = IIf(.nullable, 1, 0)
        Me.chkAutoIncrement.Value = IIf(.autoIncrement, 1, 0)
    End With
    
End Sub

Private Sub txtArchivo_GotFocus()

    markText Me.txtArchivo
    
End Sub

Private Sub txtArchivoDAO_GotFocus()

    markText Me.txtArchivoDAO
    
End Sub

Private Sub txtAtributo_GotFocus()

    markText Me.txtAtributo
    
End Sub

Private Sub txtDatabase_GotFocus()

    markText Me.txtDatabase
    
End Sub

Private Sub txtEndpoint_GotFocus()

    markText Me.txtEndpoint
    
End Sub

Private Sub txtIP_GotFocus()

    markText Me.txtIP
    
End Sub

Private Sub txtPwd_GotFocus()

    markText Me.txtPwd
    
End Sub

Private Sub txtUrl_Click()

    markText Me.txtUrl
    
End Sub

Private Sub txtUser_GotFocus()

    markText Me.txtUser
    
End Sub
