VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmIntegracionDatosSII 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integracion Datos SII"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   -240
      TabIndex        =   0
      Top             =   -600
      Width           =   12255
      Begin FlexEdGrid3.FEd3Grid Grid 
         Height          =   2295
         Left            =   480
         TabIndex        =   5
         Top             =   1680
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
         Cols            =   2
         Rows            =   2
         FixedCols       =   1
         FixedRows       =   1
         ScrollBars      =   3
         AllowUserResizing=   0
         HighLight       =   1
         SelectionMode   =   0
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   -1  'True
         Locked          =   0   'False
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1635
      End
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cancelar"
         Height          =   435
         Left            =   9360
         TabIndex        =   2
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Bt_Export 
         Caption         =   "&Integracion SII"
         Height          =   735
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "FrmIntegracionDatosSII.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nueva empresa"
         Top             =   840
         Width           =   2280
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Top             =   1020
         Width           =   660
      End
   End
End
Attribute VB_Name = "FrmIntegracionDatosSII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const C_IDEMPRESA = 0
Private Const C_RUT = 1
Private Const C_NOMCORTO = 2
Private Const C_CLAVESII = 3
Private Const C_VCLAVESII = 4
Private Const C_SELEC = 5

Private Const LASTCOL = C_SELEC

Private Sub SetupForm()

   Grid.Cols = LASTCOL + 1

   Call FGrSetup(Grid)
   
   Grid.TextMatrix(0, C_IDEMPRESA) = ""
   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_NOMCORTO) = "Nombre Corto"
   Grid.TextMatrix(0, C_CLAVESII) = "Clave SII"
   Grid.TextMatrix(0, C_VCLAVESII) = ""
   Grid.TextMatrix(0, C_SELEC) = "Seleccionar"

   Grid.ColWidth(C_IDEMPRESA) = 0
   Grid.ColWidth(C_RUT) = 1600
   Grid.ColWidth(C_NOMCORTO) = 2000
   Grid.ColWidth(C_CLAVESII) = 1500
   Grid.ColWidth(C_VCLAVESII) = 0
   Grid.ColWidth(C_SELEC) = 900


   Grid.ColAlignment(C_RUT) = flexAlignLeftCenter
   Grid.ColAlignment(C_NOMCORTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_CLAVESII) = flexAlignCenterCenter
   Grid.ColAlignment(C_SELEC) = flexAlignCenterCenter


End Sub

Private Sub Bt_Close_Click()
Unload Me
End Sub

Private Sub Bt_Export_Click()
Dim i As Integer
Dim Q1 As String
Dim j As Integer
Dim Count As Integer
      
   Count = 0
   For i = Grid.FixedRows To Grid.rows - 1
    If Grid.TextMatrix(i, C_SELEC) <> "NO" And Grid.TextMatrix(i, C_CLAVESII) <> "" Then
        If CrearNuevoAnoDelAdmin(Grid.TextMatrix(i, C_IDEMPRESA), CbItemData(Cb_Ano), vFmtCID(Grid.TextMatrix(i, C_RUT)), Grid.TextMatrix(i, C_NOMCORTO)) Then
        Call ImportarEmpresaSII(CbItemData(Cb_Ano), Grid.TextMatrix(i, C_IDEMPRESA), Grid.TextMatrix(i, C_RUT), Grid.TextMatrix(i, C_VCLAVESII))
        Count = Count + 1
        End If
    End If
   Next i
   
   If Count > 0 Then
        MsgBox "Proceso de Importacion realizado con exito Cantidad de empresas: " & Count, vbInformation, "Importacion Empresas SII"
   Else
        MsgBox "No se realizo ninguna importacion", vbInformation, "Importacion Empresas SII"
   End If
   
   #If DATACON = 1 Then
    Call CloseDb(DbMain)
    Call OpenDbAdm
   #End If
   
   Call LoadGrid


End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ano As Long

Ano = Year(Now) - 4
For i = 0 To 6
    Call CbAddItem(Cb_Ano, Ano, Ano)
    Ano = Ano + 1
Next i
Cb_Ano.ListIndex = 4

Call SetupForm
Call LoadGrid

End Sub

Private Sub LoadGrid()
Dim Q1 As String
Dim Rs As Recordset
Dim i As Integer

    Q1 = "SELECT E.IdEmpresa,E.Rut,E.NombreCorto,E.ClaveSII "
    Q1 = Q1 & " FROM Empresas AS E"
    Q1 = Q1 & " LEFT JOIN EmpresasAno AS EA ON EA.idEmpresa = E.IdEmpresa"
    Q1 = Q1 & " WHERE EA.IdEmpresa IS NULL"
    Q1 = Q1 & " ORDER BY E.ClaveSII DESC"
    Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
      'Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      
        Grid.TextMatrix(i, C_IDEMPRESA) = vFld(Rs("IdEmpresa"))
        Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")))
        Grid.TextMatrix(i, C_NOMCORTO) = vFld(Rs("NombreCorto"))
        Grid.TextMatrix(i, C_CLAVESII) = String(Len(Trim$(vFld(Rs("ClaveSII")) & "")), 42)
        Grid.TextMatrix(i, C_VCLAVESII) = vFld(Rs("ClaveSII")) ' String(Len(Trim$(vFld(Rs("ClaveSII")) & "")), 42)
        Grid.TextMatrix(i, C_SELEC) = "NO"
      
      
   Rs.MoveNext
      i = i + 1
   Loop


End Sub


Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
'Dim Gr As FlexEdGrid3
'Set Gr = Grid
'Grid.Refresh
Grid.Redraw = False


Grid.TextMatrix(Row, C_VCLAVESII) = Value
Grid.TextMatrix(Row, C_CLAVESII) = "pruebaaaa" 'String(Len(Trim$(Value & "")), 42)
Grid.Refresh
Call FGrVRows(Grid, 1)
   Grid.Row = Grid.FixedRows
   Grid.Redraw = True


End Sub
Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
If Col = C_CLAVESII Then
   Grid.TxBox.MaxLength = 30
      EdType = FEG_Edit
   End If
End Sub

Private Sub Grid_CbEditKeyPress(ByVal Col As Integer, KeyAscii As Integer)
If Col = C_CLAVESII Then
      Call KeyUpper(KeyAscii)
   End If
End Sub

Private Sub Grid_DblClick()
Select Case Grid.Col
   
      Case C_SELEC
        If Grid.TextMatrix(Grid.Row, C_CLAVESII) <> "" Then
         Grid.TextMatrix(Grid.Row, Grid.Col) = IIf(Grid.TextMatrix(Grid.Row, Grid.Col) = "SI", "NO", "SI")
        Else
            MsgBox "Para selecionar debe tener la Clave del SII"
        End If
End Select
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
If Grid.Col = C_CLAVESII Then
      'Call KeyName(KeyAscii)
   End If
End Sub



