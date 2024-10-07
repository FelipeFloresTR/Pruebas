Attribute VB_Name = "BaseImponible14ter"
Option Explicit

Public Const BASEIMP_INGRESOS = 1
Public Const BASEIMP_EGRESOS = 2
Public Const BASEIMP_TOTALES = 3

Public Const MAX_ITEMBASEIMP = 6

Public gBaseImp14Ter(BASEIMP_TOTALES, MAX_ITEMBASEIMP) As String

'Asistente c�lculo Impuesto Primera Categor�a

Public Const C_MAX_ASISTIMPPRIMCAT = 8
Public Const C_MAX_ITEMASISTPRIMCAT = 5

Public gStrAsistImpPrimCat(C_MAX_ASISTIMPPRIMCAT) As String
Public gValAsistImpPrimCat(C_MAX_ASISTIMPPRIMCAT, C_MAX_ITEMASISTPRIMCAT) As Double

Public Sub InitBaseImponible14Ter()

   gBaseImp14Ter(BASEIMP_INGRESOS, 0) = "Total de ingresos anuales percibidos en el ejercicio (y devengados en los casos que corresponda), a valor nominal"
   gBaseImp14Ter(BASEIMP_INGRESOS, 1) = "Ingresos percibidos"
   gBaseImp14Ter(BASEIMP_INGRESOS, 2) = "Ingreso diferido imputado en el ejercicio"
   gBaseImp14Ter(BASEIMP_INGRESOS, 3) = "Ingresos devengados"
   gBaseImp14Ter(BASEIMP_INGRESOS, 4) = "Participaciones e intereses percibidos"
   gBaseImp14Ter(BASEIMP_INGRESOS, 5) = "Otros ingresos percibidos o devengados"
   gBaseImp14Ter(BASEIMP_INGRESOS, 6) = "Cr�dito sobre activos fijos adquiridos y pagados en el ejercicio"

   gBaseImp14Ter(BASEIMP_EGRESOS, 0) = "Total de egresos anuales efectivamente pagados en el ejercicio, a valor nominal"
   gBaseImp14Ter(BASEIMP_EGRESOS, 1) = "Costo directo de los bienes o servicios"
   gBaseImp14Ter(BASEIMP_EGRESOS, 2) = "Remuneraciones"
   gBaseImp14Ter(BASEIMP_EGRESOS, 3) = "Adquisici�n de bienes del activo realizable y fijo"
   gBaseImp14Ter(BASEIMP_EGRESOS, 4) = "Intereses pagados"
   gBaseImp14Ter(BASEIMP_EGRESOS, 5) = "P�rdidas de ejercicios anteriores"
   gBaseImp14Ter(BASEIMP_EGRESOS, 6) = "Otros gastos deducidos de los ingresos"
   
   gBaseImp14Ter(BASEIMP_TOTALES, 0) = "Base imponible del impuesto de primera categor�a"
   gBaseImp14Ter(BASEIMP_TOTALES, 1) = "Mayor valor enajenaci�n bienes del activo fisico no depreciables, de acuerdo a la LIR"

End Sub

Public Function GetValAjustesELC(ByVal TipoAjuste As Integer, ByVal IdItemAjuste As Integer, Optional ByVal IdComp As Long = 0) As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetValAjustesELC = 0
   
   If TipoAjuste = 0 Or IdItemAjuste = 0 Then
      Exit Function
   End If
   
   Q1 = "SELECT Valor FROM AjustesExtLibCaja "
   Q1 = Q1 & " WHERE TipoAjuste = " & TipoAjuste & " AND IdItemAjuste = " & IdItemAjuste
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetValAjustesELC = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)

End Function

Public Function GetTotCta_CodF22_14Ter(ByVal CodF22_14Ter As Integer, ByVal DebCred As String, Optional ByVal IdComp As Long = 0) As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetTotCta_CodF22_14Ter = 0
   
   If CodF22_14Ter = 0 Then
      Exit Function
   End If
   
   CodF22_14Ter = HomologaCod14D(CodF22_14Ter)
   
   If DebCred = "D" Then   'd�bitos
      Q1 = "SELECT Sum(MovComprobante.Debe) FROM (MovComprobante "
   Else    ' "C"           'cr�ditos
      Q1 = "SELECT Sum(MovComprobante.Haber) FROM (MovComprobante "
   End If
   
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante")
   
   If DebCred = "D" Then   'd�bitos
      Q1 = Q1 & " WHERE Tipo = " & TC_EGRESO
   Else           ' "C"           'cr�ditos
      Q1 = Q1 & " WHERE Tipo = " & TC_INGRESO
   End If
   
   If IdComp > 0 Then
      Q1 = Q1 & " AND Comprobante.IdComp = " & IdComp
   End If
   
   Q1 = Q1 & " AND CodF22_14Ter = " & CodF22_14Ter
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetTotCta_CodF22_14Ter = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
End Function


'esta funci�n es igual a la anterior pero al rev�s porque es la Notas de Cr�dito (Adem�s selecciona s�lo los mov. comprobantes cuyo doc asociado sea NC)
Public Function GetTotCta_CodF22_14Ter_NC(ByVal CodF22_14Ter As Integer, ByVal DebCred As String) As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetTotCta_CodF22_14Ter_NC = 0
   
   If CodF22_14Ter = 0 Then
      Exit Function
   End If
   
   CodF22_14Ter = HomologaCod14D(CodF22_14Ter)
   
   If DebCred = "D" Then   'd�bitos
      Q1 = "SELECT Sum(MovComprobante.Debe) FROM (((MovComprobante "
   Else    ' "C"           'cr�ditos
      Q1 = "SELECT Sum(MovComprobante.Haber) FROM (((MovComprobante "
   End If
   
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   
   Q1 = Q1 & " INNER JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   
   If DebCred = "D" Then   'd�bitos
      Q1 = Q1 & " WHERE Tipo = " & TC_EGRESO
   Else           ' "C"           'cr�ditos
      Q1 = Q1 & " WHERE Tipo = " & TC_INGRESO
   End If
   
   Q1 = Q1 & " AND CodF22_14Ter = " & HomologaCod14D(CodF22_14Ter)
   
   If DebCred = "D" Then   'd�bitos
      Q1 = Q1 & " AND Documento.TipoLib = " & LIB_VENTAS & " AND EsRebaja <> 0 "
   
   Else    ' "C"           'cr�ditos
      Q1 = Q1 & " AND Documento.TipoLib = " & LIB_COMPRAS & " AND EsRebaja <> 0 "
   
   End If
   
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetTotCta_CodF22_14Ter_NC = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Sub InitAsistImpPrimCat()

   gStrAsistImpPrimCat(1) = "IDPC Sobre Base Imponible"
   gStrAsistImpPrimCat(2) = "Cr�ditos contra Impuesto de Primera Categor�a"
   gStrAsistImpPrimCat(3) = "Cr�dito art 33 Bis"
   gStrAsistImpPrimCat(4) = "Cr�dito asociado a ingreso diferido"
   gStrAsistImpPrimCat(5) = "Cr�dito asociado a retiros, dividendos y participaciones percibidas"
   gStrAsistImpPrimCat(6) = "IDPC a Neto a Pagar"
   gStrAsistImpPrimCat(7) = "Mayor Valor Enajenaci�n Bienes Activo F�sico no Dep."
   gStrAsistImpPrimCat(8) = "IDPC a Pagar"

End Sub

