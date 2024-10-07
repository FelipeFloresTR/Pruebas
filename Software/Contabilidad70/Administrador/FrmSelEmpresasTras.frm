VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSelEmpresasTras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Empresa y año desde a Traspasar"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "FrmSelEmpresasTras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr_Sort 
      Caption         =   "Ordenar por"
      Height          =   975
      Left            =   8340
      TabIndex        =   11
      Top             =   3240
      Width           =   1275
      Begin VB.OptionButton Op_SortNombre 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Op_SortRUT 
         Caption         =   "RUT"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.ListBox Ls_Ano 
      Height          =   1035
      Left            =   8400
      TabIndex        =   1
      Top             =   1740
      Width           =   1155
   End
   Begin VB.ListBox Ls_Empresas 
      Height          =   4935
      Left            =   1620
      TabIndex        =   0
      Top             =   720
      Width           =   6615
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8400
      TabIndex        =   3
      Top             =   960
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Default         =   -1  'True
      Height          =   315
      Left            =   8400
      TabIndex        =   2
      Top             =   600
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar PgrBar 
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   6480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Traspaso"
      Height          =   315
      Index           =   3
      Left            =   3000
      TabIndex        =   16
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Lbl_TablaPG 
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   6120
      Width           =   8535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Index           =   0
      Left            =   420
      Picture         =   "FrmSelEmpresasTras.frx":000C
      Top             =   480
      Width           =   690
   End
   Begin VB.Label La_nEmp 
      AutoSize        =   -1  'True
      Caption         =   "000"
      Height          =   195
      Left            =   8280
      TabIndex        =   10
      Top             =   5640
      Width           =   270
   End
   Begin VB.Label La_demo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "DEMO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002A01A6&
      Height          =   330
      Left            =   8520
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "(*) con datos"
      Height          =   255
      Left            =   8400
      TabIndex        =   8
      Top             =   2820
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Año Desde:"
      Height          =   195
      Index           =   2
      Left            =   8400
      TabIndex        =   7
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Corto"
      Height          =   315
      Index           =   1
      Left            =   4020
      TabIndex        =   6
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUT"
      Height          =   315
      Index           =   6
      Left            =   1620
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Corto"
      Height          =   315
      Index           =   0
      Left            =   2820
      TabIndex        =   4
      Top             =   480
      Width           =   2115
   End
End
Attribute VB_Name = "FrmSelEmpresasTras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_NOMLARGO = 2
Private Const C_IDEMPRESA = 3
Private Const C_IDPERFIL = 4
Private Const C_PRIV = 5
Private Const C_TRASPASO = 6

Private Const C_FCIERRE = 2
Private Const C_FAPERTURA = 3

Dim lRc As Integer
Dim lsEmpresa As ClsCombo
Dim LsAno As ClsCombo



Friend Function FSelect() As Integer
   Me.Show vbModal
   
   FSelect = lRc
End Function
Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub
Private Sub Bt_Sel_Click()
   Dim IdEmpresa As Long
   Dim IdEmpresaTras As Long
   Dim Ano As Integer
   Dim Rut As String
   Dim Nombre As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   Call AddDebug("FrmSelEmpresasTras.Bt_Sel_Click: 1", 1)
   
   If lsEmpresa.ListIndex < 0 Then
      Exit Sub
   End If
   
   If LsAno.ListIndex < 0 Then
      Exit Sub
   End If
   
   Call AddDebug("FrmSelEmpresasTras.Bt_Sel_Click: 2 - " & lsEmpresa.ListIndex & " - " & LsAno.ListIndex, 1)
   
   MousePointer = vbHourglass
   DoEvents
   
   IdEmpresa = Val(lsEmpresa.Matrix(C_IDEMPRESA))
   'Ano = LsAno.List(LsAno.ListIndex)
   Ano = Val(LsAno.ItemData)
   Rut = lsEmpresa.ItemData
   Nombre = lsEmpresa.List2(lsEmpresa.ListIndex)
   
   Call AddDebug("FrmSelEmpresasTras.Bt_Sel_Click: 3", 1)

   'Creo o chequeo base de datos de la empresa
'   If CrearNuevoAno(IdEmpresa, ano, Rut, Nombre) = False Then
'      MousePointer = vbDefault
'      Exit Sub
'   End If
   
   
   Call AddDebug("FrmSelEmpresasTras.Bt_Sel_Click: 4", 1)
   
   'ASIGNO DATOS A LA ESTRUCTURA
   gEmpresa.Rut = Rut
   gEmpresa.NombreCorto = Nombre
   gEmpresa.id = IdEmpresa
   gEmpresa.Ano = Ano
'   Debug.Print "gEmpresa.Ano =" & gEmpresa.Ano
   gEmpresa.FCierre = vFmt(LsAno.Matrix(C_FCIERRE))
   gEmpresa.FApertura = vFmt(LsAno.Matrix(C_FAPERTURA))
   
   If MsgBox("¿Está seguro que desea Traspasar la empresa " & Nombre & " desde el año " & Ano & "?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Empresas' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Empresas ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DbMain, Q1)
   
   
   'esto se hizo especial para la Asoc. de AFP porque tenían varias empresas con el mismo RUT.
   'El campo RutDisp sirve para imprimirlo en los membretes
   Q1 = "SELECT RutDisp, IdTras FROM Empresas WHERE IdEmpresa = " & IdEmpresa
   'Q1 = "SELECT RutDisp FROM Empresas WHERE IdEmpresa = " & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)
   gEmpresa.RutDisp = ""
   If Not Rs.EOF Then
      gEmpresa.RutDisp = FwDecrypt1(vFld(Rs("RutDisp")), KEY_CRYP + 10)
      IdEmpresaTras = vFld(Rs("IdTras"))
   End If
   Call CloseRs(Rs)
   
   gUsuario.idPerfil = lsEmpresa.Matrix(C_IDPERFIL)
   gUsuario.Priv = lsEmpresa.Matrix(C_PRIV)
   
   
   'esto es para la importacion desde Access a SQL
   ' 3114594 FPR
   Dim ultimoAno As Boolean
   Dim AnoCurso As Integer
   Dim PathDbAnoAnt As String
   Dim ConnStr As String
   'Dim DbAnoCurso As Database
   Dim DbAnoAnt As Database
   
   ultimoAno = False
   AnoCurso = Ano
   
   Do While ultimoAno = False
        
        PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & AnoCurso & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
        'PathDbAnoAnt = Replace(gDbPath & "\Empresas\" & AnoCurso & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", "")
        
        
        If ExistFile(PathDbAnoAnt) Then
             '3391062 se comenta ya que se debe conectar con la bd lpcontab PASSW_LEXCONT
             'ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
             ConnStr = ";PWD=" & PASSW_LEXCONT & ";"
              
             Set DbAnoAnt = OpenDatabase(Replace(gDbPath, "LPContabSQL", "LPContab") & "\" & BD_COMUN, False, False, ConnStr)
             'Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
             Call TrasLpContab(DbMain, DbAnoAnt, gEmpresa.id, AnoCurso, IdEmpresaTras, PgrBar, Lbl_TablaPG)
             
             'ffv 3391062
             ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
             
             'Set DbAnoAnt = Nothing
             Call CloseDb(DbAnoAnt)
             
             Call OpenDbEmpresa(DbAnoAnt, PathDbAnoAnt)
           
             Call LinkMdbAdm(DbAnoAnt)
             
             Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
             
             Call TrasLPEmpresa(DbMain, DbAnoAnt, gEmpresa.id, AnoCurso, IdEmpresaTras, PgrBar, Lbl_TablaPG)
             Call ModCtaParamEmpresa(DbMain, gEmpresa.id, AnoCurso)
             
        Else
         ultimoAno = True
            
        End If
        AnoCurso = AnoCurso + 1
        
   Loop
   
   ' FIN 3114594
      
   lRc = vbOK
   
   MsgBox "Empresa Traspasada correctamente ", vbInformation
   Call FillList
   'Unload Me
   
   
   Call AddDebug("FrmSelEmpresasTras.Bt_Sel_Click: FIN", 1)
   
   If gDbPathAux <> "" Then
    gDbPath = gDbPathAux
   End If
   Me.MousePointer = vbDefault
   
End Sub


Private Sub Form_Load()


'   If MsgBox("¿Está seguro que desea borrar este perfil?", vbQuestion + vbYesNo) = vbNo Then
'      Unload Me
'   End If
   If DbMain Is Nothing Then
       Call OpenMsSql
   End If

   lRc = vbCancel
   PgrBar.Value = 0
   
   Call AddDebug("FrmSelEmpresasTras: Load", 1)
   
   Set LsAno = New ClsCombo
   Call LsAno.SetControl(Ls_Ano)
   
   Call AddDebug("FrmSelEmpresasTras: Después de ClsCombo", 1)
   
   'If gVarIniFile.SelEmprPorRUT Then
      Op_SortRUT = True    'llama a FillList
   'Else
      'Op_SortNombre = True    'llama a FillList
   'End If
   
   Call AddDebug("FrmSelEmpresasTras: Después de FillList", 1)
   
   La_demo.Visible = gAppCode.Demo
   'Fr_Sort.Visible = Not gAppCode.Demo
   
End Sub

Private Sub FillList()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Wh As String
   
      
   Set lsEmpresa = New ClsCombo
   Call lsEmpresa.SetControl(Ls_Empresas)
      
   If gUsuario.Nombre = gAdmUser Then
'      Q1 = "SELECT Empresas.idEmpresa, Rut, NombreCorto, 0 as idPerfil, 0 as Privilegios"
'      Q1 = Q1 & " FROM Empresas"
'      Q1 = Q1 & " WHERE Estado = 0 "   'Empresas activas
      
      Q1 = "SELECT Empresas.idEmpresa, Empresas.Rut, Empresas.NombreCorto, 0 as idPerfil, 0 as Privilegios, IIF(EmpresasAno.idEmpresa IS NULL,0,1) AS Traspaso"
      Q1 = Q1 & " FROM Empresas"
      Q1 = Q1 & " LEFT JOIN EmpresasAno ON EmpresasAno.idEmpresa = Empresas.IdEmpresa"
      Q1 = Q1 & " Where Estado = 0" 'Empresas activas
      Q1 = Q1 & " AND Empresas.IdTras IS NOT NULL"
      Q1 = Q1 & " GROUP BY Empresas.idEmpresa, Empresas.Rut, Empresas.NombreCorto,EmpresasAno.idEmpresa"
      
      If gAppCode.Demo Then
         Q1 = Q1 & " AND RUT IN ('1','2','3')"
      End If
   Else
      Q1 = "SELECT Empresas.idEmpresa, Rut, NombreCorto, UsuarioEmpresa.idPerfil, Perfiles.Privilegios"
      Q1 = Q1 & " FROM (Empresas INNER JOIN UsuarioEmpresa ON Empresas.idEmpresa=UsuarioEmpresa.IdEmpresa)"
      Q1 = Q1 & " LEFT JOIN Perfiles ON UsuarioEmpresa.idPerfil = Perfiles.idPerfil"
      Q1 = Q1 & " WHERE UsuarioEmpresa.idUsuario = " & gUsuario.IdUsuario
      Q1 = Q1 & " AND Empresas.Estado = 0 "   'Empresas activas

      If gAppCode.Demo Then
         Q1 = Q1 & " AND RUT IN ('1','2','3')"
      End If
   
   End If
   
   If Op_SortNombre Then
      Q1 = Q1 & " ORDER BY NombreCorto, RUT"
   Else
'      Q1 = Q1 & " ORDER BY right('0' & RUT,8)"

      Q1 = Q1 & " ORDER BY right(" & SqlConcat(gDbType, "'00000000000'", "Rut") & ", 12)"  ' 19 mar 2019 - pam: falla el orden porque algunos tienen puntos y otros no
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      lsEmpresa.AddItem FmtStRut(vFld(Rs("Rut"))) & vbTab & "   " & IIf(vFld(Rs("Traspaso")) = 0, "NO", "SI") & vbTab & "       " & vFld(Rs("NombreCorto"))
      lsEmpresa.ItemData(lsEmpresa.NewIndex) = vFld(Rs("Rut"))
      lsEmpresa.Matrix(C_NOMLARGO, lsEmpresa.NewIndex) = vFld(Rs("NombreCorto"))
      lsEmpresa.Matrix(C_IDEMPRESA, lsEmpresa.NewIndex) = vFld(Rs("idEmpresa"))
      lsEmpresa.Matrix(C_IDPERFIL, lsEmpresa.NewIndex) = vFld(Rs("idPerfil"))
      lsEmpresa.Matrix(C_TRASPASO, lsEmpresa.NewIndex) = vFld(Rs("Traspaso"))
      
      If gUsuario.Nombre = gAdmUser Then
         lsEmpresa.Matrix(C_PRIV, lsEmpresa.NewIndex) = PRV_ADMIN
      Else
         lsEmpresa.Matrix(C_PRIV, lsEmpresa.NewIndex) = vFld(Rs("Privilegios"))
      End If
      
      Rs.MoveNext
      
      If gAppCode.Demo And lsEmpresa.ListCount >= 3 Then
         Exit Do
      End If
      
   Loop
   Call CloseRs(Rs)
   
   La_nEmp = Format(lsEmpresa.ListCount, NUMFMT)
   
'   If gAppCode.NivProd = VER_5EMP And lsEmpresa.ListCount > 5 Then
   If lsEmpresa.ListCount > gMaxEmpLicencia And Not gAppCode.Demo And gDbType = SQL_ACCESS Then
      MsgBox1 "Esta versión sólo permite trabajar con a lo más " & gMaxEmpLicencia & " empresas." & vbCrLf & vbCrLf & "Utilice el administrador para eliminar algunas empresas y así poder utilizar el sistema.", vbExclamation
      Bt_Sel.Enabled = False
   End If
   
End Sub



Private Sub Ls_Ano_DblClick()
   Call PostClick(Bt_Sel)
   
   Call AddDebug("FrmSelEmpresasTras.Ls_Ano_DblClick: FIN", 1)

End Sub

Private Sub Ls_Empresas_Click_old()
   Dim Q1 As String
   Dim Ano As Integer
   Dim Rs As Recordset
   Dim UltAño As Integer
   Dim i As Integer
   
   'Set LsAno = New ClsCombo
   'Call LsAno.SetControl(Ls_Ano)
   
   LsAno.Clear
   Ano = Year(Int(Now))
   
   Q1 = "SELECT Max(Ano) as MaxAno FROM EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
      
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
            
      If vFld(Rs("MaxAno")) > 0 Then
          
         'agregamos un año más del último, para que lo pueda crear, si lo desea
         LsAno.AddItem vFld(Rs("MaxAno")) + 1
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.ListIndex = 0
            
         Q1 = "SELECT Ano, FCierre, FApertura FROM EmpresasAno"
         Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
         Q1 = Q1 & " ORDER BY Ano DESC"
         Call LsAno.FillCombo(DbMain, Q1, -1)
         
         UltAño = Val(LsAno.list(Ls_Ano.ListCount - 1))
         
         For i = 1 To 4
            LsAno.AddItem UltAño - i
            LsAno.ItemData(LsAno.NewIndex) = 0
         Next i
                
         If LsAno.ListCount > 1 Then
            LsAno.ListIndex = 1  'seleccionamos último año existente
         End If
   
      Else
         'No existe ningun año creado, ofrecemos: año-4, año-3, año-2, año-1, año y año+1
         LsAno.AddItem Ano + 1
         LsAno.ItemData(LsAno.NewIndex) = 0
         
         LsAno.AddItem Ano
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.ListIndex = LsAno.NewIndex
         
         LsAno.AddItem Ano - 1
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.AddItem Ano - 2
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.AddItem Ano - 3
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.AddItem Ano - 4
         LsAno.ItemData(LsAno.NewIndex) = 0
               
      End If
      
   End If
   
   Call CloseRs(Rs)
   
End Sub

Private Sub Ls_Empresas_Click()
   Dim Q1 As String
   Dim Ano As Integer, MaxAno As Integer
   Dim Rs As Recordset
   Dim UltAño As Integer
   Dim i As Integer
   Dim AnoTope As Integer
      
   Call AddDebug("FrmSelEmpresasTras.Ls_Empresas_Click: 1", 1)
      
   LsAno.Clear
   Ano = Year(Int(Now))
   
   Q1 = "SELECT Max(Ano) as MaxAno FROM EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
      
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      MaxAno = vFld(Rs("MaxAno"))
   End If
   Call CloseRs(Rs)
   
   If MaxAno <= 0 Then
      MaxAno = Year(Now)
   End If
   
   Call AddDebug("FrmSelEmpresasTras.Ls_Empresas_Click: 2 - " & MaxAno, 1)
   
   Q1 = "SELECT Ano, FCierre, FApertura FROM EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
   Q1 = Q1 & " ORDER BY Ano DESC"
   Set Rs = OpenRs(DbMain, Q1)
   
   Call AddDebug("FrmSelEmpresasTras.Ls_Empresas_Click: 3", 1)
   
   AnoTope = Year(Now) + 2
   
   For i = AnoTope To 2000 Step -1
      Call LsAno.AddItem(i, i, 0, 0)
      If i = MaxAno Then
         LsAno.ListIndex = LsAno.NewIndex
      End If
   
      Do Until Rs.EOF
         Ano = vFld(Rs("Ano"))
         
         If Ano > AnoTope Then ' años muy del futuro o tiene mal la fecha del computador
            Exit Do
         End If
         
         If i = Ano Then
            LsAno.list(LsAno.NewIndex) = Ano & " *"
            LsAno.Matrix(C_FCIERRE, LsAno.NewIndex) = vFld(Rs("FCierre"))
            LsAno.Matrix(C_FAPERTURA, LsAno.NewIndex) = vFld(Rs("FApertura"))
            Rs.MoveNext
            Exit Do
         Else ' If i > Ano Then 19 oct 2017 - pam: para evitar que entre en un loop infinito
            Exit Do
         End If
      Loop
   
   Next i
   
   Call CloseRs(Rs)
   
   Call AddDebug("FrmSelEmpresasTras.Ls_Empresas_Click: FIN", 1)

End Sub
Private Sub Ls_Empresas_DblClick()

   Call PostClick(Bt_Sel)
   
End Sub

Private Sub Op_SortNombre_Click()

   Me.MousePointer = vbHourglass
   Call FillList
   Call SetIniString(gIniFile, "Opciones", "SelEmprPorRut", "0")
   'gVarIniFile.SelEmprPorRUT = 0
   Me.MousePointer = vbDefault

End Sub

Private Sub Op_SortRUT_Click()

   Me.MousePointer = vbHourglass
   Call FillList
   Call SetIniString(gIniFile, "Opciones", "SelEmprPorRut", "1")
   'gVarIniFile.SelEmprPorRUT = 1
   Me.MousePointer = vbDefault

End Sub

Private Sub LinkMdbAdm(DbMain As Database, Optional ByVal bForce As Boolean = 0)
   Dim DbComun As String
   Dim ConnStr As String
   Dim Tm As Double


   
   Tm = CDbl(Now)
   DbComun = gDbPath & "\" & BD_COMUN
   
   'ConnStr = "PWD=" & PASSW_LEXCONT & ";"
   'ConnStr = Mid(gComunConnStr, 2)  ' sin el ; del inicio
   
   If bForce = False Then
      bForce = Val(GetIniString(gIniFile, "Config", "ReLink", "0"))
   End If

   ConnStr = gComunConnStr
   
   Call LinkMdbTableAdmin(DbMain, DbComun, "CodActiv", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Empresas", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "EmpresasAno", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Equivalencia", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Impuestos", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Monedas", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Param", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "PlanAvanzado", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "PlanBasico", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "PlanIntermedio", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Regiones", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Timbraje", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "TipoValor", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Usuarios", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "UsuarioEmpresa", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "Perfiles", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "IPC", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "TipoDocs", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "ControlEmpresa", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "PcUsr", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "PlanCuentasSII", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "FactorActAnual", , bForce, , ConnStr, True)
   Call LinkMdbTableAdmin(DbMain, DbComun, "CapPropioSimplAnual", , bForce, , ConnStr, True)
  
   If gFunciones.IFRS Then    'ya no se usa
      Call LinkMdbTableAdmin(DbMain, DbComun, "IFRS_PlanIFRS", , bForce, , ConnStr, True)
   End If
   
   'gLinkF22 = LinkDbfTable(DbMain, gHRPath & "\PAR", "NContrib.dbf", "HR_NContrib", "FoxPro 2.0", , False)
'   If Not gLinkF22 Then
'      gLinkF22 = ExistFile(gHRPath & "\PAR\BD_HR_admin.mdb")
'   End If
   
'   gPathForm22 = "\FORM22"
'   gPathPlan22 = "\PLAN22"
   
   
   If gFunciones.RazFinancieras Then
      'Call LinkMdbTableAdmin(DbMain, DbComun, "CuentasRazon", , , , ConnStr)    'ya no se linkea porque está en la DB de la empresa
      Call LinkMdbTableAdmin(DbMain, DbComun, "RazonesFin", , bForce, , ConnStr)
   End If
   
'   If gFunciones.ExpFUT Then
'      gLinkParFUT = LinkDbfTable(DbMain, gHRPath & "\PAR", "HFTPAR52.dbf", "HR_FUTGrItems", "FoxPro 2.0", , False)
'   End If
   
   Debug.Print "LinkMdbAdm: Tiempo: " & Format((CDbl(Now) - Tm) / TimeSerial(0, 0, 1), NUMFMT) & " [s]"
   

   
End Sub


Private Function LinkMdbTableAdmin(Db As Database, ByVal MdbPath As String, ByVal TableName As String, Optional ByVal NewName As String = "", Optional bForce As Boolean = 0, Optional bMsg As Boolean = 1, Optional ConnString As String = "", Optional ByVal bForceIfNotLinked As Boolean = False) As Boolean
   Dim Tbl As TableDef
   Dim i As Integer
   Dim ConnStr As String, Msg As String, Conn1 As String, TConnect As String, TPWD As String
   Dim Q1 As String
   
   On Error Resume Next
   
   LinkMdbTableAdmin = False
   
   If Trim(NewName) = "" Then
      NewName = TableName
   End If
      
   If Left(ConnString, 1) = ";" Then
      ConnString = Mid(ConnString, 2)
   End If
      
   'FCA: se agrega verificación de bForce para que lo haga de todas maneras, si se requiere
      
   If Db.TableDefs(NewName).Connect = "" And Not bForceIfNotLinked Then ' No es una tabla linkeada, se perderian datos
      If Err = 0 Then
         LinkMdbTableAdmin = False
         Exit Function
      End If
      Err.Clear

   End If
   
   Set Tbl = Db.TableDefs(NewName)

   If Not Tbl Is Nothing Then

      If bForce = False Then
      
         TConnect = Tbl.Connect
         TPWD = "PWD=" & GetTxConnectInfo(TConnect, "PWD") & ";"
      
         ' Si estaba linkeado sin clave pero ahora la base tiene clave, y la clave es diferente => FORCE
         If ConnString <> "" And StrComp(TPWD, ConnString, vbTextCompare) <> 0 Then
            bForce = True
         Else
         
            ' El Count sirve para verificar si está linkeada con la password correcta
            
            If Err = 0 And StrComp(GetTxConnectInfo(TConnect, "DATABASE"), AbsPath(MdbPath), vbTextCompare) = 0 _
               And StrComp(Tbl.SourceTableName, TableName, vbTextCompare) = 0 Then
               
               If SameMdb(GetTxConnectInfo(TConnect, "DATABASE"), MdbPath, True) Then ' por si cambiaron la unidad y la base sigue existiendo  z:\datos\lpremu.mdb
'               If ExistFile(MdbPath) Then
                  LinkMdbTableAdmin = True
                  Exit Function
               End If
               
            End If
            
         End If
         
      End If
            
   End If
   
   Debug.Print "LinkMdbTableAdmin: relinkeando " & NewName & " - psw=" & (ConnString <> "")
   
   ConnStr = ";DATABASE=" & AbsPath(MdbPath) & ";" & "PWD=Fw#420!&+;" 'ConnString

'se cambia dbmain por db gcb201021
   Q1 = "DROP TABLE " & NewName
   Call ExecSQL(Db, Q1, W.InDesign)

   Err.Clear
   Db.TableDefs.Delete NewName   ' Si ya existía, la eliminamos
   If Err.Number <> 0 And Err.Number <> 3265 And Err.Number <> 3011 Then  ' si no existe, todo bien
      Msg = "Error al eliminar la tabla vinculada '" & NewName & "'."

      If bMsg Then
         MsgErr Msg
      End If
   
      Call AddLog("LinkMdb: " & Msg & " Err=" & Err & ", " & Error)
      LinkMdbTableAdmin = False
      
      Exit Function
   End If
   
   Err.Clear
   
   Set Tbl = New TableDef
   Tbl.Connect = ConnStr
   Tbl.SourceTableName = TableName
   Tbl.Name = NewName
   
   Db.TableDefs.Append Tbl
   Db.TableDefs.Refresh
   
   LinkMdbTableAdmin = (Err = 0)
   
   If Err Then
      Msg = "Error al vincular la tabla '" & TableName & "' ubicada en " & MdbPath & "."

      If bMsg Then
         MsgErr Msg
      End If
      
      Debug.Print "FALLÓ LinkMdb(" & TableName & ") Err=" & Err & ", " & Error
      
      Call AddLog("LinkMdb: " & Msg & " Err=" & Err & ", " & Error)
   
   End If
      
End Function

Private Function OpenDbEmpresa(DbEmp As Database, Path As String, Optional ByVal Rut As String = "", Optional ByVal Ano As Integer = 0) As Integer

   Dim DbName As String
   Dim Passw As String, SqlErr As String
   
   On Error Resume Next
   
   OpenDbEmpresa = True
          
   If Path <> "" Then
    DbName = Path
   Else
    If Ano > 0 Then
       If Rut <> "" Then
          DbName = gDbPath & "\Empresas\" & Ano & "\" & Rut & ".mdb"
       Else
          DbName = gDbPath & "\Empresas\" & Ano & "\" & gEmpresa.Rut & ".mdb"
       End If
       
    ElseIf Rut <> "" Then
       DbName = gDbPath & "\Empresas\" & gEmpresa.Ano & "\" & Rut & ".mdb"
    Else
       DbName = gDbPath & "\Empresas\" & gEmpresa.Ano & "\" & gEmpresa.Rut & ".mdb"
    End If
   End If

'   If Rut <> "" Then
'      Passw = PASSW_PREFIX & Rut
'   Else
'      Passw = PASSW_PREFIX & gEmpresa.Rut
'   End If
   
    '2868088
   If Rut <> "" Then
      Passw = PASSW_PREFIX & Rut
   Else
      Passw = PASSW_PREFIX & gEmpresa.Rut
   End If
   'FIN 2868088
   
   
   Call AddLog("OpenDbEmpresa: DbName:[" & DbName & "]", 2)
   
   'Call SetDbSecurity(DbName, Passw, gCfgFile, SG_SEGCFG, gEmpresa.ConnStr)
   Call SetDbSecurityAdm(DbName, PASSW_PREFIX & gEmpresa.Rut, Replace(gCfgFile, "Administrador", "HyperContabilidad"), "FW6T9R54WX3A", gEmpresa.ConnStr)

   Err.Clear
   'Set DbEmp = OpenDatabase(DbName, True, False, ConnStr) ' MODO EXCLUSIVO
   Set DbEmp = OpenDatabase(DbName, False, False, gEmpresa.ConnStr)
   'gEmpresa.ConnStr = Mid(gEmpresa.ConnStr, 2) 'sin el ; del principio   FCA: 2 feb 2016 se comenta esta línea
   
   If Err Then
      SqlErr = "Error " & Err & ", '" & Error & "'"
   
      If Err = 3356 Then
         MsgBox1 "Ya existe algún usuario trabajando con la empresa seleccionada.", vbExclamation
         OpenDbEmpresa = False
      End If
   
   End If
   
   If (Err Or DbEmp Is Nothing) And Err <> 3356 Then
      MsgBox SqlErr & vbCrLf & DbName, vbExclamation
      OpenDbEmpresa = False
   End If
   
   Call ChkDbSize(DbMain, 200 * 1024) ' 200 MB
   
   On Error GoTo 0
   
   Call AddLog("OpenDbEmpresa: fin OK", 2)

End Function
Private Sub ChkDbSize(Db As Database, ByVal MAXKBSIZE As Long)
   Dim FSize As Long
   
   If SqlType(Db, Db.Connect) <> SQL_ACCESS Then
      Exit Sub
   End If

   FSize = FileLen(Db.Name)

   If FSize / 1024 > MAXKBSIZE Then
      MsgBox1 "ATENCIÓN" & vbCrLf & vbCrLf & "El tamaño de la base de datos supera los " & Format(MAXKBSIZE / 1024, NUMFMT) & " MBytes," & vbCrLf & "es necesario que utilice la opción para compactarla.", vbExclamation
   End If

End Sub

Public Sub SetDbSecurityAdm(ByVal DbPath As String, ByVal Passw As String, ByVal CfgFile As String, ByVal SegCfg As String, ConnStr As String)
   Dim Db As Database
   Dim Cfg As String, ConnStr1 As String
   Dim Seg As Boolean
   
   On Error Resume Next
   
   Cfg = GetIniString(CfgFile, "Config", "Secur", "")
   'MsgBox "ruta : " & CfgFile & "clave CFG: " & Cfg & " CLAVE SegCfg: " & SegCfg
   If Cfg <> SegCfg Then
      Seg = True
      ConnStr1 = ""
      ConnStr = ";PWD=" & Passw & ";"
   Else
      Seg = False
      ConnStr1 = ";PWD=" & Passw & ";"
      ConnStr = ""
   End If

   Err.Clear

   ' Probamos a abrir la base con lo contrario de la seguridad esperada
   Set Db = OpenDatabase(DbPath, True, False, ConnStr1)
   If Not Db Is Nothing Then ' si pudo abrir, entonces hay que cambiar
      
      If Seg Then
         Db.NewPassword "", Passw   ' le pone clave
      Else
         Db.NewPassword Passw, ""   ' le quita la clave
      End If
      
      If Err Then
         Call AddLog("SetDbSecurity: Seg=" & Seg & ", " & DbPath & ", Error " & Err & ", " & Err.Description)
      End If
      
      Call CloseDb(Db)
   Else
      Call AddLog("SetDbSecurity: No se pudo quitar/poner clave a la base de datos, " & DbPath & ", Error " & Err & ", " & Err.Description)
      Debug.Print "*** No se pudo quitar/poner clave a la base de datos."
   End If
   
End Sub



