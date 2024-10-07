VERSION 5.00
Begin VB.Form FrmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Respaldos"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   Icon            =   "FrmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_Msg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FrmBackup.frx":000C
      Top             =   780
      Width           =   10515
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   9840
      Top             =   180
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   180
      Picture         =   "FrmBackup.frx":0012
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "� Respald� su informaci�n esta semana ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   Image2 = Image1

   Tx_Msg = vbCrLf & "                  * * *  IMPORTANTE  * * *"
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   
   If gDbType = SQL_ACCESS Then
      Tx_Msg = Tx_Msg & "Carpeta a respaldar: " & W.AppPath
      Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   End If
   
   Tx_Msg = Tx_Msg & "Es de suma importancia realizar RESPALDOS de la informaci�n en forma peri�dica. Es responsabilidad del usuario o empresa definir una pol�tica adecuada al respecto."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "En el caso de la p�rdida de informaci�n debido al ataque de un virus, la falla de un disco, el da�o de la base de datos, etc., la �nica forma de recuperar y no perder el trabajo de meses, es recurrir a los respaldos."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Los programas pueden ser instalados nuevamente, pero si no hay respaldos, la informaci�n ingresada se perder� irremediablemente."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Los respaldos se deben almacenar en MEDIOS EXTERNOS. No deben ser guardados en el mismo disco o en el mismo equipo en que se encuentra la aplicaci�n. Un virus puede destruir el contenido de todo el disco o los discos conectados al equipo, o bien puede fallar el disco en que se hizo el respaldo."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Es importante verificar que los respaldos queden bien hechos, de modo que cuando se necesiten puedan ser utilizados. Cada cierto tiempo se debe recuperar un respaldo y verificar si la informaci�n es la correcta."
   Tx_Msg = Tx_Msg & " Los dispositivos donde se hace el respaldo (ej discos externos, pendrives, DVDs, CDs, etc.), es recomendable que se almacenen fuera de la oficina."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   If gDbType = SQL_ACCESS Then
       Tx_Msg = Tx_Msg & "Para nuestra aplicaci�n " & App.Title & ", usted deber�a respaldar toda la carpeta '" & W.AppPath & "'."
       Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
       
    '   If gDbType = SQL_MYSQL Then
    '      Tx_Msg = Tx_Msg & "En esta versi�n del programa, la base de datos est� en un servidor MySQL, debe solicitar la asistencia de un t�cnico para realizar el respado de los datos."
    '      Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
    '   End If
       
       Tx_Msg = Tx_Msg & "Si el respaldo lo hace hoy, cree en la unidad una carpeta llamada 'Respaldo_" & Format(Now, "yymmdd") & "'. En esta carpeta agregue el contenido de la carpeta '" & W.AppPath & "' y toda otra informaci�n importante para usted."
       Tx_Msg = Tx_Msg & vbCrLf
       Tx_Msg = Tx_Msg & "En el siguiente respaldo, cree una carpeta con la nueva fecha y agregue en esta nueva carpeta su informaci�n, ojal� en una unidad diferente."
       Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
       Tx_Msg = Tx_Msg & "Se recomienda tener tres o m�s unidades de respaldo distintas e ir rotando su uso. De ese modo si se da�a una unidad, quedan las otras."
       Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Else
       Tx_Msg = Tx_Msg & "Para nuestra aplicaci�n " & App.Title & " SQL, usted deber� usar el programa Microsoft SQL Server Management Studio para efectuar el respectivo respaldo de informaci�n."
       Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   
   End If
   Tx_Msg = Tx_Msg & "Recuerde mantener actualizado su Antivirus y chequear peri�dicamente sus discos para reducir los riesgos."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   
End Sub

