Attribute VB_Name = "LpContab_ANSi_Textos"
Option Explicit

'Esta funci�n se desarroll� para corregir el error que gener� el primer Script de la base de datos
'que qued� grabado en UTF8 en vez de ANSI
Function CorrigeTextosConAcentosScriptUTF8()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Txt As String, Txt2 As String, Txt3 As String, Txt4 As String
   Dim bUpdate As Boolean
   
   'CodActiv
   Q1 = "SELECT Codigo, Descrip, Version FROM CodActiv"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descrip"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE CodActiv SET Descrip = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE Codigo = '" & vFld(Rs("Codigo")) & "' AND Version = " & vFld(Rs("Version"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'CT_ComprobanteBase
   Q1 = "SELECT IdComp, Nombre, Descrip, Glosa FROM CT_ComprobanteBase "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Nombre"))
      Txt2 = vFld(Rs("Descrip"))
      Txt3 = vFld(Rs("Glosa"))
      
      bUpdate = CorrigeTexto(Txt) Or CorrigeTexto(Txt2) Or CorrigeTexto(Txt3)
      If bUpdate Then
         Q1 = "UPDATE CT_ComprobanteBase SET Nombre = '" & ParaSQL(Txt) & "', Descrip = '" & ParaSQL(Txt2) & "', Glosa = '" & ParaSQL(Txt3) & "'"
         Q1 = Q1 & " WHERE IdComp = " & vFld(Rs("IdComp"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'CT_MovComprobanteBase
   Q1 = "SELECT IdMov, IdComp, Glosa FROM CT_MovComprobanteBase"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Glosa"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE CT_MovComprobanteBase SET Glosa = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdMov = '" & vFld(Rs("IdMov")) & "' AND IdComp = " & vFld(Rs("IdComp"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   
   Loop
   
   Call CloseRs(Rs)
   
   'CT_Comprobante
   Q1 = "SELECT IdComp, Nombre, Descrip, Glosa FROM CT_Comprobante "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Nombre"))
      Txt2 = vFld(Rs("Descrip"))
      Txt3 = vFld(Rs("Glosa"))
      
      bUpdate = CorrigeTexto(Txt) Or CorrigeTexto(Txt2) Or CorrigeTexto(Txt3)
      If bUpdate Then
         Q1 = "UPDATE CT_Comprobante SET Nombre = '" & ParaSQL(Txt) & "', Descrip = '" & ParaSQL(Txt2) & "', Glosa = '" & ParaSQL(Txt3) & "'"
         Q1 = Q1 & " WHERE IdComp = " & vFld(Rs("IdComp"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'CT_MovComprobante
   Q1 = "SELECT IdMov, IdComp, Glosa FROM CT_MovComprobante"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Glosa"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE CT_MovComprobante SET Glosa = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdMov = '" & vFld(Rs("IdMov")) & "' AND IdComp = " & vFld(Rs("IdComp"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'IFRS_PlanIFRS
   Q1 = "SELECT IdCuenta, Descripcion FROM IFRS_PlanIFRS"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE IFRS_PlanIFRS SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'PlanBasico
   Q1 = "SELECT IdCuenta, Descripcion FROM PlanBasico"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE PlanBasico SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'PlanIntermedio
   Q1 = "SELECT IdCuenta, Descripcion FROM PlanIntermedio"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE PlanIntermedio SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)

   'PlanAvanzado
   Q1 = "SELECT IdCuenta, Descripcion FROM PlanAvanzado"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE PlanAvanzado SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'Cuentas
   Q1 = "SELECT IdCuenta, Descripcion FROM Cuentas"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE Cuentas SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'PlanCuentasSII
   Q1 = "SELECT IdPlanCuentasSII, DescripSII FROM PlanCuentasSII"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("DescripSII"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE PlanCuentasSII SET DescripSII = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdPlanCuentasSII = " & vFld(Rs("IdPlanCuentasSII"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'RazonesFin
   Q1 = "SELECT IdRazon, Nombre, TxtNumerador, TxtDenominador FROM RazonesFin "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Nombre"))
      Txt2 = vFld(Rs("TxtNumerador"))
      Txt3 = vFld(Rs("TxtDenominador"))
      
      bUpdate = CorrigeTexto(Txt) Or CorrigeTexto(Txt2) Or CorrigeTexto(Txt3)
      If bUpdate Then
         Q1 = "UPDATE RazonesFin SET Nombre = '" & ParaSQL(Txt) & "', TxtNumerador = '" & ParaSQL(Txt2) & "', TxtDenominador = '" & ParaSQL(Txt3) & "'"
         Q1 = Q1 & " WHERE IdRazon = " & vFld(Rs("IdRazon"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'Regiones
   Q1 = "SELECT Id, Comuna FROM Regiones"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Comuna"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE Regiones SET Comuna = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE Id = " & vFld(Rs("Id"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'TipoDocs
   Q1 = "SELECT Id, Nombre FROM TipoDocs"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Nombre"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE TipoDocs SET Nombre = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE Id = " & vFld(Rs("Id"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'TipoValor
   Q1 = "SELECT IdTValor, Valor, Tit1, Tit2, TitCompleto FROM TipoValor "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Valor"))
      Txt2 = vFld(Rs("Tit1"))
      Txt3 = vFld(Rs("Tit2"))
      Txt4 = vFld(Rs("TitCompleto"))
      
      bUpdate = CorrigeTexto(Txt) Or CorrigeTexto(Txt2) Or CorrigeTexto(Txt3)
      If bUpdate Then
         Q1 = "UPDATE TipoValor SET Valor = '" & ParaSQL(Txt) & "', Tit1 = '" & ParaSQL(Txt2) & "', Tit2 = '" & ParaSQL(Txt3) & "', TitCompleto = '" & ParaSQL(Txt4) & "'"
         Q1 = Q1 & " WHERE IdTValor = " & vFld(Rs("IdTValor"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   

   MsgBox1 "Proceso finalizado", vbOKOnly + vbInformation
   
End Function



Function CorrigeTexto(Buf As String) As Boolean
    If InStr(Buf, "?") <= 0 Then
       Exit Function
    End If

   Buf = Replace(Buf, "ma?z", "ma�z")
   Buf = Replace(Buf, "tub?rculos", "tub�rculos")
   Buf = Replace(Buf, "ca?a", "ca�a")
   Buf = Replace(Buf, "az?car", "az�car")
   Buf = Replace(Buf, "producci?n", "producci�n")
   Buf = Replace(Buf, "c?tricos", "c�tricos")
   Buf = Replace(Buf, "?rboles", "�rboles")
   Buf = Replace(Buf, "caf?", "caf�")
   Buf = Replace(Buf, "t?", "t�")
   Buf = Replace(Buf, "arom?ticas", "arom�ticas")
   Buf = Replace(Buf, "farmac?uticas", "farmac�uticas")
   Buf = Replace(Buf, "Cr?a", "Cr�a")
   Buf = Replace(Buf, "vicu?as", "vicu�as")
   Buf = Replace(Buf, "cam?lidos", "cam�lidos")
   Buf = Replace(Buf, "agr?colas", "agr�colas")
   Buf = Replace(Buf, "combinaci?n", "combinaci�n")
   Buf = Replace(Buf, "cr?a", "cr�a")
   Buf = Replace(Buf, "explotaci?n", "explotaci�n")
   Buf = Replace(Buf, "ganader?a", "ganader�a")
   Buf = Replace(Buf, "propagaci?n", "propagaci�n")
   Buf = Replace(Buf, "Explotaci?n", "Explotaci�n")
   Buf = Replace(Buf, "Extracci?n", "Extracci�n")
   Buf = Replace(Buf, "Recolecci?n", "Recolecci�n")
   Buf = Replace(Buf, "forestaci?n", "forestaci�n")
   Buf = Replace(Buf, "retribuci?n", "retribuci�n")
   Buf = Replace(Buf, "extinci?n", "extinci�n")
   Buf = Replace(Buf, "prevenci?n", "prevenci�n")
   Buf = Replace(Buf, "mar?tima", "mar�tima")
   Buf = Replace(Buf, "factor?a", "factor�a")
   Buf = Replace(Buf, "extracci?n", "extracci�n")
   Buf = Replace(Buf, "reproducci?n", "reproducci�n")
   Buf = Replace(Buf, "Reproducci?n", "Reproducci�n")
   Buf = Replace(Buf, "crust?ceos", "crust�ceos")
   Buf = Replace(Buf, "carb?n", "carb�n")
   Buf = Replace(Buf, "petr?leo", "petr�leo")
   Buf = Replace(Buf, "metal?feros", "metal�feros")
   Buf = Replace(Buf, "fabricaci?n", "fabricaci�n")
   Buf = Replace(Buf, "qu?micos", "qu�micos")
   Buf = Replace(Buf, "Elaboraci?n", "Elaboraci�n")
   Buf = Replace(Buf, "conservaci?n", "conservaci�n")
   Buf = Replace(Buf, "c?rnicos", "c�rnicos")
   Buf = Replace(Buf, "Producci?n", "Producci�n")
   Buf = Replace(Buf, "salm?nidos", "salm�nidos")
   Buf = Replace(Buf, "acu?ticos", "acu�ticos")
   Buf = Replace(Buf, "elaboraci?n", "elaboraci�n")
   Buf = Replace(Buf, "l?cteos", "l�cteos")
   Buf = Replace(Buf, "s?mola", "s�mola")
   Buf = Replace(Buf, "gr?nulos", "gr�nulos")
   Buf = Replace(Buf, "moliner?a", "moliner�a")
   Buf = Replace(Buf, "almid?n", "almid�n")
   Buf = Replace(Buf, "panader?a", "panader�a")
   Buf = Replace(Buf, "pasteler?a", "pasteler�a")
   Buf = Replace(Buf, "confiter?a", "confiter�a")
   Buf = Replace(Buf, "farin?ceos", "farin�ceos")
   Buf = Replace(Buf, "informaci?n", "informaci�n")
   Buf = Replace(Buf, "Destilaci?n", "Destilaci�n")
   Buf = Replace(Buf, "rectificaci?n", "rectificaci�n")
   Buf = Replace(Buf, "alcoh?licas", "alcoh�licas")
   Buf = Replace(Buf, "Preparaci?n", "Preparaci�n")
   Buf = Replace(Buf, "Fabricaci?n", "Fabricaci�n")
   Buf = Replace(Buf, "art?culos", "art�culos")
   Buf = Replace(Buf, "te?ido", "te�ido")
   Buf = Replace(Buf, "talabarter?a", "talabarter�a")
   Buf = Replace(Buf, "guarnicioner?a", "guarnicioner�a")
   Buf = Replace(Buf, "carpinter?a", "carpinter�a")
   Buf = Replace(Buf, "cart?n", "cart�n")
   Buf = Replace(Buf, "Impresi?n", "Impresi�n")
   Buf = Replace(Buf, "impresi?n", "impresi�n")
   Buf = Replace(Buf, "refinaci?n", "refinaci�n")
   Buf = Replace(Buf, "qu?micas", "qu�micas")
   Buf = Replace(Buf, "b?sicas", "b�sicas")
   Buf = Replace(Buf, "nitr?geno", "nitr�geno")
   Buf = Replace(Buf, "pl?sticos", "pl�sticos")
   Buf = Replace(Buf, "sint?tico", "sint�tico")
   Buf = Replace(Buf, "pirot?cnicos", "pirot�cnicos")
   Buf = Replace(Buf, "farmac?uticos", "farmac�uticos")
   Buf = Replace(Buf, "bot?nicos", "bot�nicos")
   Buf = Replace(Buf, "c?maras", "c�maras")
   Buf = Replace(Buf, "renovaci?n", "renovaci�n")
   Buf = Replace(Buf, "pl?stico", "pl�stico")
   Buf = Replace(Buf, "construcci?n", "construcci�n")
   Buf = Replace(Buf, "cer?mica", "cer�mica")
   Buf = Replace(Buf, "hormig?n", "hormig�n")
   Buf = Replace(Buf, "met?licos", "met�licos")
   Buf = Replace(Buf, "Fundici?n", "Fundici�n")
   Buf = Replace(Buf, "dep?sitos", "dep�sitos")
   Buf = Replace(Buf, "calefacci?n", "calefacci�n")
   Buf = Replace(Buf, "cuchiller?a", "cuchiller�a")
   Buf = Replace(Buf, "ferreter?a", "ferreter�a")
   Buf = Replace(Buf, "electr?nicos", "electr�nicos")
   Buf = Replace(Buf, "perif?rico", "perif�rico")
   Buf = Replace(Buf, "medici?n", "medici�n")
   Buf = Replace(Buf, "navegaci?n", "navegaci�n")
   Buf = Replace(Buf, "irradiaci?n", "irradiaci�n")
   Buf = Replace(Buf, "electr?nico", "electr�nico")
   Buf = Replace(Buf, "m?dico", "m�dico")
   Buf = Replace(Buf, "terap?utico", "terap�utico")
   Buf = Replace(Buf, "?pticos", "�pticos")
   Buf = Replace(Buf, "fotogr?fico", "fotogr�fico")
   Buf = Replace(Buf, "magn?ticos", "magn�ticos")
   Buf = Replace(Buf, "el?ctricos", "el�ctricos")
   Buf = Replace(Buf, "distribuci?n", "distribuci�n")
   Buf = Replace(Buf, "bater?as", "bater�as")
   Buf = Replace(Buf, "?ptica", "�ptica")
   Buf = Replace(Buf, "el?ctrico", "el�ctrico")
   Buf = Replace(Buf, "iluminaci?n", "iluminaci�n")
   Buf = Replace(Buf, "dom?stico", "dom�stico")
   Buf = Replace(Buf, "veh?culos", "veh�culos")
   Buf = Replace(Buf, "propulsi?n", "propulsi�n")
   Buf = Replace(Buf, "v?lvulas", "v�lvulas")
   Buf = Replace(Buf, "transmisi?n", "transmisi�n")
   Buf = Replace(Buf, "elevaci?n", "elevaci�n")
   Buf = Replace(Buf, "manipulaci?n", "manipulaci�n")
   Buf = Replace(Buf, "conformaci?n", "conformaci�n")
   Buf = Replace(Buf, "m?quinas", "m�quinas")
   Buf = Replace(Buf, "metal?rgica", "metal�rgica")
   Buf = Replace(Buf, "carrocer?as", "carrocer�as")
   Buf = Replace(Buf, "Construcci?n", "Construcci�n")
   Buf = Replace(Buf, "bisuter?a", "bisuter�a")
   Buf = Replace(Buf, "m?dicos", "m�dicos")
   Buf = Replace(Buf, "oftalmol?gicos", "oftalmol�gicos")
   Buf = Replace(Buf, "odontol?gicos", "odontol�gicos")
   Buf = Replace(Buf, "Reparaci?n", "Reparaci�n")
   Buf = Replace(Buf, "miner?a", "miner�a")
   Buf = Replace(Buf, "reparaci?n", "reparaci�n")
   Buf = Replace(Buf, "dom?sticos", "dom�sticos")
   Buf = Replace(Buf, "Instalaci?n", "Instalaci�n")
   Buf = Replace(Buf, "Generaci?n", "Generaci�n")
   Buf = Replace(Buf, "energ?a", "energ�a")
   Buf = Replace(Buf, "el?ctrica", "el�ctrica")
   Buf = Replace(Buf, "hidroel?ctricas", "hidroel�ctricas")
   Buf = Replace(Buf, "termoel?ctricas", "termoel�ctricas")
   Buf = Replace(Buf, "Transmisi?n", "Transmisi�n")
   Buf = Replace(Buf, "Distribuci?n", "Distribuci�n")
   Buf = Replace(Buf, "Regasificaci?n", "Regasificaci�n")
   Buf = Replace(Buf, "tuber?a", "tuber�a")
   Buf = Replace(Buf, "regasificaci?n", "regasificaci�n")
   Buf = Replace(Buf, "Captaci?n", "Captaci�n")
   Buf = Replace(Buf, "Evacuaci?n", "Evacuaci�n")
   Buf = Replace(Buf, "eliminaci?n", "eliminaci�n")
   Buf = Replace(Buf, "Recuperaci?n", "Recuperaci�n")
   Buf = Replace(Buf, "descontaminaci?n", "descontaminaci�n")
   Buf = Replace(Buf, "gesti?n", "gesti�n")
   Buf = Replace(Buf, "l?neas", "l�neas")
   Buf = Replace(Buf, "p?blico", "p�blico")
   Buf = Replace(Buf, "ingenier?a", "ingenier�a")
   Buf = Replace(Buf, "Demolici?n", "Demolici�n")
   Buf = Replace(Buf, "el?ctricas", "el�ctricas")
   Buf = Replace(Buf, "gasfiter?a", "gasfiter�a")
   Buf = Replace(Buf, "Terminaci?n", "Terminaci�n")
   Buf = Replace(Buf, "perfumer?a", "perfumer�a")
   Buf = Replace(Buf, "cosm?ticos", "cosm�ticos")
   Buf = Replace(Buf, "papeler?a", "papeler�a")
   Buf = Replace(Buf, "cient?ficos", "cient�ficos")
   Buf = Replace(Buf, "quir?rgicos", "quir�rgicos")
   Buf = Replace(Buf, "inform?ticos", "inform�ticos")
   Buf = Replace(Buf, "s?lidos", "s�lidos")
   Buf = Replace(Buf, "l?quidos", "l�quidos")
   Buf = Replace(Buf, "peque?os", "peque�os")
   Buf = Replace(Buf, "verduler?as", "verduler�as")
   Buf = Replace(Buf, "botiller?as", "botiller�as")
   Buf = Replace(Buf, "m?sica", "m�sica")
   Buf = Replace(Buf, "ortop?dicos", "ortop�dicos")
   Buf = Replace(Buf, "joyer?a", "joyer�a")
   Buf = Replace(Buf, "relojer?a", "relojer�a")
   Buf = Replace(Buf, "le?a", "le�a")
   Buf = Replace(Buf, "artesan?as", "artesan�as")
   Buf = Replace(Buf, "antig?edades", "antig�edades")
   Buf = Replace(Buf, "v?a", "v�a")
   Buf = Replace(Buf, "telef?nica", "telef�nica")
   Buf = Replace(Buf, "locomoci?n", "locomoci�n")
   Buf = Replace(Buf, "tuber?as", "tuber�as")
   Buf = Replace(Buf, "mar?timo", "mar�timo")
   Buf = Replace(Buf, "v?as", "v�as")
   Buf = Replace(Buf, "a?rea", "a�rea")
   Buf = Replace(Buf, "frigor?ficos", "frigor�ficos")
   Buf = Replace(Buf, "dep?sito", "dep�sito")
   Buf = Replace(Buf, "parqu?metros", "parqu�metros")
   Buf = Replace(Buf, "acu?tico", "acu�tico")
   Buf = Replace(Buf, "a?reo", "a�reo")
   Buf = Replace(Buf, "Manipulaci?n", "Manipulaci�n")
   Buf = Replace(Buf, "mensajer?a", "mensajer�a")
   Buf = Replace(Buf, "m?vil", "m�vil")
   Buf = Replace(Buf, "banqueter?a", "banqueter�a")
   Buf = Replace(Buf, "concesi?n", "concesi�n")
   Buf = Replace(Buf, "alimentaci?n", "alimentaci�n")
   Buf = Replace(Buf, "Edici?n", "Edici�n")
   Buf = Replace(Buf, "peri?dicas", "peri�dicas")
   Buf = Replace(Buf, "edici?n", "edici�n")
   Buf = Replace(Buf, "pel?culas", "pel�culas")
   Buf = Replace(Buf, "cinematogr?ficas", "cinematogr�ficas")
   Buf = Replace(Buf, "televisi?n", "televisi�n")
   Buf = Replace(Buf, "postproducci?n", "postproducci�n")
   Buf = Replace(Buf, "exhibici?n", "exhibici�n")
   Buf = Replace(Buf, "grabaci?n", "grabaci�n")
   Buf = Replace(Buf, "Programaci?n", "Programaci�n")
   Buf = Replace(Buf, "Telefon?a", "Telefon�a")
   Buf = Replace(Buf, "Televisi?n", "Televisi�n")
   Buf = Replace(Buf, "al?mbricas", "al�mbricas")
   Buf = Replace(Buf, "m?viles", "m�viles")
   Buf = Replace(Buf, "inal?mbrica", "inal�mbrica")
   Buf = Replace(Buf, "inal?mbricas", "inal�mbricas")
   Buf = Replace(Buf, "sat?lite", "sat�lite")
   Buf = Replace(Buf, "programaci?n", "programaci�n")
   Buf = Replace(Buf, "inform?tica", "inform�tica")
   Buf = Replace(Buf, "consultor?a", "consultor�a")
   Buf = Replace(Buf, "inform?ticas", "inform�ticas")
   Buf = Replace(Buf, "tecnolog?a", "tecnolog�a")
   Buf = Replace(Buf, "intermediaci?n", "intermediaci�n")
   Buf = Replace(Buf, "inversi?n", "inversi�n")
   Buf = Replace(Buf, "cr?dito", "cr�dito")
   Buf = Replace(Buf, "compensaci?n", "compensaci�n")
   Buf = Replace(Buf, "Administraci?n", "Administraci�n")
   Buf = Replace(Buf, "asesor?a", "asesor�a")
   Buf = Replace(Buf, "Evaluaci?n", "Evaluaci�n")
   Buf = Replace(Buf, "da?os", "da�os")
   Buf = Replace(Buf, "representaci?n", "representaci�n")
   Buf = Replace(Buf, "jur?dica", "jur�dica")
   Buf = Replace(Buf, "ra?ces", "ra�ces")
   Buf = Replace(Buf, "s?ndicos", "s�ndicos")
   Buf = Replace(Buf, "jur?dicas", "jur�dicas")
   Buf = Replace(Buf, "tenedur?a", "tenedur�a")
   Buf = Replace(Buf, "auditor?a", "auditor�a")
   Buf = Replace(Buf, "dise?o", "dise�o")
   Buf = Replace(Buf, "t?cnica", "t�cnica")
   Buf = Replace(Buf, "revisi?n", "revisi�n")
   Buf = Replace(Buf, "an?lisis", "an�lisis")
   Buf = Replace(Buf, "t?cnicos", "t�cnicos")
   Buf = Replace(Buf, "opini?n", "opini�n")
   Buf = Replace(Buf, "p?blica", "p�blica")
   Buf = Replace(Buf, "decoraci?n", "decoraci�n")
   Buf = Replace(Buf, "ampliaci?n", "ampliaci�n")
   Buf = Replace(Buf, "fotograf?as", "fotograf�as")
   Buf = Replace(Buf, "fotograf?a", "fotograf�a")
   Buf = Replace(Buf, "Asesor?a", "Asesor�a")
   Buf = Replace(Buf, "peque?as", "peque�as")
   Buf = Replace(Buf, "traducci?n", "traducci�n")
   Buf = Replace(Buf, "interpretaci?n", "interpretaci�n")
   Buf = Replace(Buf, "p?blicas", "p�blicas")
   Buf = Replace(Buf, "cient?ficas", "cient�ficas")
   Buf = Replace(Buf, "t?cnicas", "t�cnicas")
   Buf = Replace(Buf, "cl?nicas", "cl�nicas")
   Buf = Replace(Buf, "dotaci?n", "dotaci�n")
   Buf = Replace(Buf, "tur?sticos", "tur�sticos")
   Buf = Replace(Buf, "cerrajer?a", "cerrajer�a")
   Buf = Replace(Buf, "investigaci?n", "investigaci�n")
   Buf = Replace(Buf, "Desratizaci?n", "Desratizaci�n")
   Buf = Replace(Buf, "desinfecci?n", "desinfecci�n")
   Buf = Replace(Buf, "jardiner?a", "jardiner�a")
   Buf = Replace(Buf, "preparaci?n", "preparaci�n")
   Buf = Replace(Buf, "Organizaci?n", "Organizaci�n")
   Buf = Replace(Buf, "calificaci?n", "calificaci�n")
   Buf = Replace(Buf, "administraci?n", "administraci�n")
   Buf = Replace(Buf, "Regulaci?n", "Regulaci�n")
   Buf = Replace(Buf, "facilitaci?n", "facilitaci�n")
   Buf = Replace(Buf, "econ?mica", "econ�mica")
   Buf = Replace(Buf, "Previsi?n", "Previsi�n")
   Buf = Replace(Buf, "afiliaci?n", "afiliaci�n")
   Buf = Replace(Buf, "Ense?anza", "Ense�anza")
   Buf = Replace(Buf, "cient?fico", "cient�fico")
   Buf = Replace(Buf, "t?cnico", "t�cnico")
   Buf = Replace(Buf, "formaci?n", "formaci�n")
   Buf = Replace(Buf, "educaci?n", "educaci�n")
   Buf = Replace(Buf, "ense?anza", "ense�anza")
   Buf = Replace(Buf, "atenci?n", "atenci�n")
   Buf = Replace(Buf, "odontol?gica", "odontol�gica")
   Buf = Replace(Buf, "odont?logos", "odont�logos")
   Buf = Replace(Buf, "cl?nicos", "cl�nicos")
   Buf = Replace(Buf, "enfermer?a", "enfermer�a")
   Buf = Replace(Buf, "toxic?manos", "toxic�manos")
   Buf = Replace(Buf, "f?sica", "f�sica")
   Buf = Replace(Buf, "espect?culos", "espect�culos")
   Buf = Replace(Buf, "esc?nicas", "esc�nicas")
   Buf = Replace(Buf, "art?sticas", "art�sticas")
   Buf = Replace(Buf, "compa??as", "compa��as")
   Buf = Replace(Buf, "m?sicos", "m�sicos")
   Buf = Replace(Buf, "hist?ricos", "hist�ricos")
   Buf = Replace(Buf, "zool?gicos", "zool�gicos")
   Buf = Replace(Buf, "Hip?dromos", "Hip�dromos")
   Buf = Replace(Buf, "Gesti?n", "Gesti�n")
   Buf = Replace(Buf, "f?tbol", "f�tbol")
   Buf = Replace(Buf, "Promoci?n", "Promoci�n")
   Buf = Replace(Buf, "organizaci?n", "organizaci�n")
   Buf = Replace(Buf, "tem?ticos", "tem�ticos")
   Buf = Replace(Buf, "pol?ticas", "pol�ticas")
   Buf = Replace(Buf, "tel?fonos", "tel�fonos")
   Buf = Replace(Buf, "Peluquer?a", "Peluquer�a")
   Buf = Replace(Buf, "guarder?a", "guarder�a")
   Buf = Replace(Buf, "peluquer?a", "peluquer�a")
   Buf = Replace(Buf, "ba?os", "ba�os")
   Buf = Replace(Buf, "p?blicos", "p�blicos")
   Buf = Replace(Buf, "?rganos", "�rganos")
   Buf = Replace(Buf, "Seg?n", "Seg�n")
   Buf = Replace(Buf, "Centralizaci?n", "Centralizaci�n")
   Buf = Replace(Buf, "Correcci?n", "Correcci�n")
   Buf = Replace(Buf, "Contabilizaci?n", "Contabilizaci�n")
   Buf = Replace(Buf, "Declaraci?n", "Declaraci�n")
   Buf = Replace(Buf, "A?o", "A�o")
   Buf = Replace(Buf, "Depreciaci?n", "Depreciaci�n")
   Buf = Replace(Buf, "B?sicos", "B�sicos")
   Buf = Replace(Buf, "a?o", "a�o")
   Buf = Replace(Buf, "Dep?sitos", "Dep�sitos")
   Buf = Replace(Buf, "Tr?nsito", "Tr�nsito")
   Buf = Replace(Buf, "Burs?til", "Burs�til")
   Buf = Replace(Buf, "inter?s", "inter�s")
   Buf = Replace(Buf, "Cr?dito", "Cr�dito")
   Buf = Replace(Buf, "Estimaci?n", "Estimaci�n")
   Buf = Replace(Buf, "Gratificaci?n", "Gratificaci�n")
   Buf = Replace(Buf, "Pr?stamos", "Pr�stamos")
   Buf = Replace(Buf, "Garant?a", "Garant�a")
   Buf = Replace(Buf, "Mercader?as", "Mercader�as")
   Buf = Replace(Buf, "Provisi?n", "Provisi�n")
   Buf = Replace(Buf, "Biol?gicos", "Biol�gicos")
   Buf = Replace(Buf, "biol?gicos", "biol�gicos")
   Buf = Replace(Buf, "Cr?ditos", "Cr�ditos")
   Buf = Replace(Buf, "m?todo", "m�todo")
   Buf = Replace(Buf, "participaci?n", "participaci�n")
   Buf = Replace(Buf, "L?neas", "L�neas")
   Buf = Replace(Buf, "Telef?nicas", "Telef�nicas")
   Buf = Replace(Buf, "Veh?culos", "Veh�culos")
   Buf = Replace(Buf, "Autom?viles", "Autom�viles")
   Buf = Replace(Buf, "Biol?gico", "Biol�gico")
   Buf = Replace(Buf, "Inversi?n", "Inversi�n")
   Buf = Replace(Buf, "Porci?n", "Porci�n")
   Buf = Replace(Buf, "?nico", "�nico")
   Buf = Replace(Buf, "Cesant?a", "Cesant�a")
   Buf = Replace(Buf, "Participaci?n", "Participaci�n")
   Buf = Replace(Buf, "D?bito", "D�bito")
   Buf = Replace(Buf, "A?os", "A�os")
   Buf = Replace(Buf, "disposici?n", "disposici�n")
   Buf = Replace(Buf, "P?rdidas", "P�rdidas")
   Buf = Replace(Buf, "P?rdida", "P�rdida")
   Buf = Replace(Buf, "Emisi?n", "Emisi�n")
   Buf = Replace(Buf, "conversi?n", "conversi�n")
   Buf = Replace(Buf, "p?rdidas", "p�rdidas")
   Buf = Replace(Buf, "Prestaci?n", "Prestaci�n")
   Buf = Replace(Buf, "Asignaci?n", "Asignaci�n")
   Buf = Replace(Buf, "Movilizaci?n", "Movilizaci�n")
   Buf = Replace(Buf, "Colaci?n", "Colaci�n")
   Buf = Replace(Buf, "Vi?ticos", "Vi�ticos")
   Buf = Replace(Buf, "Capacitaci?n", "Capacitaci�n")
   Buf = Replace(Buf, "Mantenci?n", "Mantenci�n")
   Buf = Replace(Buf, "Remodelaci?n", "Remodelaci�n")
   Buf = Replace(Buf, "T?cnicos", "T�cnicos")
   Buf = Replace(Buf, "Calefacci?n", "Calefacci�n")
   Buf = Replace(Buf, "Tel?fono", "Tel�fono")
   Buf = Replace(Buf, "Cafeter?a", "Cafeter�a")
   Buf = Replace(Buf, "Suscripci?n", "Suscripci�n")
   Buf = Replace(Buf, "Art?culos", "Art�culos")
   Buf = Replace(Buf, "Papeler?a", "Papeler�a")
   Buf = Replace(Buf, "Enajenaci?n", "Enajenaci�n")
   Buf = Replace(Buf, "p?rdida", "p�rdida")
   Buf = Replace(Buf, "Categor?a", "Categor�a")
   Buf = Replace(Buf, "D?lar", "D�lar")
   Buf = Replace(Buf, "Consolidaci?n", "Consolidaci�n")
   Buf = Replace(Buf, "Retasaci?n", "Retasaci�n")
   Buf = Replace(Buf, "T?cnica", "T�cnica")
   Buf = Replace(Buf, "Retazaci?n", "Retazaci�n")
   Buf = Replace(Buf, "Relaci?n", "Relaci�n")
   Buf = Replace(Buf, "Amortizaci?n", "Amortizaci�n")
   Buf = Replace(Buf, "Ejecuci?n", "Ejecuci�n")
   Buf = Replace(Buf, "pr?stamos", "pr�stamos")
   Buf = Replace(Buf, "Investigaci?n", "Investigaci�n")
   Buf = Replace(Buf, "Indemnizaci?n", "Indemnizaci�n")
   Buf = Replace(Buf, "revalorizaci?n", "revalorizaci�n")
   Buf = Replace(Buf, "Aplicaci?n", "Aplicaci�n")
   Buf = Replace(Buf, "Realizaci?n", "Realizaci�n")
   Buf = Replace(Buf, "enajenaci?n", "enajenaci�n")
   Buf = Replace(Buf, "Raz?n", "Raz�n")
   Buf = Replace(Buf, "?cida", "�cida")
   Buf = Replace(Buf, "Tesorer?a", "Tesorer�a")
   Buf = Replace(Buf, "M?rgen", "M�rgen")
   Buf = Replace(Buf, "L?quida", "L�quida")
   Buf = Replace(Buf, "Rotaci?n", "Rotaci�n")
   Buf = Replace(Buf, "D?as", "D�as")
   Buf = Replace(Buf, "CAMI?A", "CAMI�A")
   Buf = Replace(Buf, "CHA?ARAL", "CHA�ARAL")
   Buf = Replace(Buf, "VICU?A", "VICU�A")
   Buf = Replace(Buf, "VI?A", "VI�A")
   Buf = Replace(Buf, "DO?IHUE", "DO�IHUE")
   Buf = Replace(Buf, "HUALA?E", "HUALA�E")
   Buf = Replace(Buf, "CA?ETE", "CA�ETE")
   Buf = Replace(Buf, "IBA?EZ", "IBA�EZ")
   Buf = Replace(Buf, "PE?AFLOR", "PE�AFLOR")
   Buf = Replace(Buf, "?U?OA", "�U�OA")
   Buf = Replace(Buf, "PE?ALOLEN", "PE�ALOLEN")
   Buf = Replace(Buf, "Liquidaci?n", "Liquidaci�n")
   Buf = Replace(Buf, "Devoluci?n", "Devoluci�n")
   Buf = Replace(Buf, "Exportaci?n", "Exportaci�n")
   Buf = Replace(Buf, "M?quina", "M�quina")
   Buf = Replace(Buf, "Electr?nico", "Electr�nico")
   Buf = Replace(Buf, "Remuneraci?n", "Remuneraci�n")
   Buf = Replace(Buf, "Champa?a", "Champa�a")
   Buf = Replace(Buf, "Alcoh?licas", "Alcoh�licas")
   Buf = Replace(Buf, "Analcoh?licas", "Analcoh�licas")
   Buf = Replace(Buf, "D?b", "D�b")
   Buf = Replace(Buf, "Cr?d", "Cr�d")
   Buf = Replace(Buf, "Espec?fico", "Espec�fico")
   Buf = Replace(Buf, "Az?car", "Az�car")
   Buf = Replace(Buf, "Hidrobiol?gicas", "Hidrobiol�gicas")
   Buf = Replace(Buf, "Comercializaci?n", "Comercializaci�n")
   Buf = Replace(Buf, "Petr?leo", "Petr�leo")
   Buf = Replace(Buf, "analcoh?licas", "analcoh�licas")
   Buf = Replace(Buf, "Retenci?n", "Retenci�n")
   Buf = Replace(Buf, "m?rgen", "m�rgen")
   Buf = Replace(Buf, "comercializaci?n", "comercializaci�n")
   Buf = Replace(Buf, "Analcoh?l", "Analcoh�l")
    CorrigeTexto = True
End Function
