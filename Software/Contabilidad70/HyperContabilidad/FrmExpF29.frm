VERSION 5.00
Begin VB.Form FrmExpF29 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar a HR-IVA F29"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "FrmExpF29.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Per�odo"
      Height          =   975
      Left            =   1260
      TabIndex        =   3
      Top             =   300
      Width           =   4395
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A�o:"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   5
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   1500
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Exp 
      Caption         =   "Exportar"
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Top             =   1500
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   300
      Picture         =   "FrmExpF29.frx":000C
      Top             =   420
      Width           =   585
   End
End
Attribute VB_Name = "FrmExpF29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_Exp_Click()
   Dim Mes As Integer
   Dim Msg As String

   Mes = ItemData(Cb_Mes)
   If Mes < 1 Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   Msg = "Recuerde que debe tener instalada la �ltima versi�n de HR-IVA F29 para el correcto traspaso de la informaci�n."
   
   If MsgBox1(Msg, vbQuestion + vbOKCancel + vbDefaultButton2) = vbCancel Then
      Exit Sub
   End If
   
'   Msg = "�ATENCION!" & vbLf & "Esta exportaci�n reemplazar� los valores actuales en el producto HR-IVA Est�ndar."
'   Msg = Msg & vbLf & vbLf & "Antes de realizar la exportaci�n, aseg�rese que ning�n usuario tenga abierto el producto HR-IVA Est�ndar con la empresa y per�odo exportado."
   
   Msg = "Empresa: " & gEmpresa.RazonSocial & vbLf & "Mes: " & gNomMes(Mes) & " " & gEmpresa.Ano
   Msg = Msg & vbLf & vbLf & "� Desea continuar ?"
   If MsgBox1(Msg, vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If

   MousePointer = vbHourglass
   Bt_Exp.Enabled = False
   DoEvents
   
'   If ExportF29(Mes) = 0 Then
   If GenDB_F29(Mes) Then
   
      Msg = "�ATENCI�N!" & vbLf & "Para terminar el proceso, ahora debe abrir el producto HR-IVA F29 y realizar el Rec�lculo."
      Msg = Msg & vbLf & vbLf & "Empresa: " & gEmpresa.RazonSocial & vbLf & "Mes: " & gNomMes(Mes) & " " & gEmpresa.Ano
      MsgBox1 Msg, vbInformation
   
   Else
      MsgBox1 "No se pudo realizar la exportaci�n.", vbExclamation
   
   End If
   
   
   Bt_Exp.Enabled = True
   MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   Call FillMes(Cb_Mes, GetMesActual())

   Tx_Ano = gEmpresa.Ano

End Sub
