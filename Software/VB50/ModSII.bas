Attribute VB_Name = "ModSII"
' M�dulo para obtener datos desde el sitio del SII
Option Explicit

'Public Const URL_FWDOC = "https://www.fairware.cl/DocsRemu.asp"
Public Const URL_FWDOC = "https://servicioslp.thomsonreuters.cl/DocsRemu.asp"

Public Const SERR_OK = 0

Public Const SERR_PGNOTFND = 404
Public Const SERR_BADPARAM = 2000
Public Const SERR_NOINFO = 2001     ' No existe la seccion con los datos
Public Const SERR_NODATA = 2002     ' No hay datos

Public Type SII_UF_t
   UF    As Double
   Buf   As Boolean
End Type

Public Type SII_Fact_t
   Fact    As Double
   FactR   As Double       ' 2 abr 2020: valor real, cuando es menor que 1, en Fact se asigna 1
   bFact   As Boolean
End Type

Public Type SII_IPC_t
   VarIpc   As Double      ' Variacion IPC mensual
   bVarIpc  As Boolean
   
   VarAcum  As Double      ' Variacion IPC Acumulada
   bVarAcum As Boolean
   
   PIpc     As Double      ' Puntos de IPC
   bPIpc    As Boolean
   
   UTM      As Long        ' UTM
   bUTM     As Boolean

End Type


' 12 sep 2017: el SII cambia la p�gina y el formato de la p�gina de las UFs
Public Function SII_GetUFs(ByVal AnoMes As Long, UFs() As SII_UF_t) As Integer ' 19 ene 2018: cambia de boolean a integer
   Dim Path As String, Page As String, T As Long, d As Integer, m As Integer, n As Integer, Buf As String
   Dim Fila As String, sUF As String, Ano As Integer, Url As String

   SII_GetUFs = 0
   Ano = AnoMes \ 100

   Url = FwWebReadPage(URL_FWDOC & "?d=UF&a=" & AnoMes \ 100 & "&u=1", "") ' u=1 para que no salte

   If Left(Url, 6) <> "##URL=" Then
      MsgBox1 "Error en la conexi�n a Internet." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline)." & vbCrLf & "Intente actualizar nuevamente.", vbExclamation
      Exit Function
   End If

   Url = Mid(Url, 7)

'   Path = "/valores_y_fechas/uf/uf" & AnoMes \ 100 & ".htm"
'   Page = FwWebReadPage("www.sii.cl", Path)
   
   Page = FwWebReadPage(Url, "")
   If Page = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
      SII_GetUFs = SERR_PGNOTFND
      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
      Exit Function
   End If
   
   m = AnoMes Mod 100
   
   Buf = "<h2>" & gNomMes(m) & "</h2>" ' 5 oct 2017: el SII cambi� de <h3> a <h2>
   T = InStr(Page, Buf)
   If T <= 0 Then
      SII_GetUFs = SERR_NOINFO
      MsgBox1 "No existe la p�gina con los datos de UF de a�o " & Ano, vbExclamation
      Exit Function
   End If
   
   Page = ReplaceStr(Page, "<br />", "<br>") ' 31 mar 2014: se agrega porque se marea con /> en medio del tr
   Page = ReplaceStr(Page, "<br/>", "<br>") ' 31 mar 2014
   
   T = T + Len(Buf) + 1

   n = 0

   For d = 1 To 11
      If d <= 10 Then
         UFs(d).Buf = 0
         UFs(d + 10).UF = 0
         UFs(d).UF = 0
         UFs(d + 10).Buf = 0
      End If
      UFs(d + 20).UF = 0
      UFs(d + 20).Buf = 0
      
      Fila = FwGetXmlTag(Page, "tr", d, T, -1)
      If Fila <> "" Then
         sUF = FwGetXmlTag(Fila, "th", IIf(d < 11, 1, 3), 0, -1)
         If sUF <> "" And sUF <> "&nbsp;" Then
            sUF = ReplaceStr(ReplaceStr(sUF, "<strong>", ""), "</strong>", "")
            If Val(sUF) <> IIf(d < 11, d, 31) Then ' verificamos que sea  d - d+10 - d+20
               Exit For
            End If
         
            If d <= 10 Then
               sUF = FwGetXmlTag(Fila, "td", 1, 0, -1)
               If sUF <> "" Then
                  UFs(d).UF = Val(ReplaceStr(ReplaceStr(sUF, ".", ""), ",", "."))
                  UFs(d).Buf = 1
                  n = n + 1
               End If
               
               sUF = FwGetXmlTag(Fila, "td", 2, 0, -1)
               If sUF <> "" Then
                  UFs(d + 10).UF = Val(ReplaceStr(ReplaceStr(sUF, ".", ""), ",", "."))
                  UFs(d + 10).Buf = 1
                  n = n + 1
               End If
            End If
            
            sUF = FwGetXmlTag(Fila, "td", 3, 0, -1)
            If sUF <> "" Then
               UFs(d + 20).UF = Val(ReplaceStr(ReplaceStr(sUF, ".", ""), ",", "."))
               UFs(d + 20).Buf = 1
               n = n + 1
            End If
                     
         End If
      End If
   
   Next d

   If n <= 0 Then
      SII_GetUFs = SERR_NODATA
      MsgBox1 "No se encontraron datos de UF para el a�o " & Ano, vbExclamation
      Exit Function
   End If

   SII_GetUFs = SERR_OK

End Function

'Public Function SII_GetUFs(ByVal AnoMes As Long, UFs() As Double) As Boolean
'   Dim Path As String, Page As String, t As Long, d As Integer, m As Integer, n As Integer
'   Dim Fila As String, sUF As String
'
'   SII_GetUFs = False
'
'   Path = "/pagina/valores/uf/uf" & AnoMes \ 100 & ".htm"
'
'   Page = FwWebReadPage("www.sii.cl", Path)
'   If Page = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
'      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
'      Exit Function
'   End If
'
'   t = InStr(Page, "<div id=""contenido"" ")
'   If t <= 0 Then
'      Exit Function
'   End If
'
'   Page = ReplaceStr(Page, "<br />", "<br>") ' 31 mar 2014: se agrega porque se marea con /> en medio del tr
'   Page = ReplaceStr(Page, "<br/>", "<br>") ' 31 mar 2014
'
'   t = InStr(t, Page, "<table ", vbTextCompare)
'
'   m = AnoMes Mod 100
'   n = 0
'
'   For d = 1 + 1 To 31 + 1 ' el 1 son los nombres de los meses
'      UFs(d - 1) = 0
'      Fila = FwGetXmlTag(Page, "tr", d, t, -1)
'      If Fila <> "" Then
'         sUF = FwGetXmlTag(Fila, "td", m, 0, -1)
'         If sUF <> "" And sUF <> "&nbsp;" Then
'            UFs(d - 1) = Val(ReplaceStr(ReplaceStr(sUF, ".", ""), ",", "."))
'            n = n + 1
'         End If
'      End If
'   Next d
'
'   SII_GetUFs = (n > 0)
'
'End Function
'Public Function SII_Factores_old(ByVal Ano As Long, Fact() As Double) As Boolean
'   Dim Path As String, Page As String, t As Long, m As Integer, n As Integer
'   Dim Fila As String, sFact As String, Td As String
'
'   SII_Factores_old = False
'
'   If Ano < 2000 Then
'      Exit Function
'   End If
'
'   Path = "/pagina/renta/" & (Ano + 1) & "/grandes_contribuyentes.htm"
'
'   Page = FwWebReadPage("www.sii.cl", Path)
'   If Page = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
'      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
'      Exit Function
'   End If
'
'   Page = ReplaceStr(Page, "   ", " ") ' 12 ene 2015: para eliminar el problema que cambien la cantidad de blancos
'   Page = ReplaceStr(Page, "  ", " ")
'
''   t = InStr(Page, " 6.- Porcentajes y Factores de Actualizaci")
'   t = InStr(1, Page, " 6.- Porcentajes y Factores de Actualizaci", vbTextCompare) ' 12 ene 2015: porque cambiaron a mayusculas
'   If t <= 0 Then
'      Exit Function
'   End If
'
'   t = InStr(t, Page, "<table ", vbTextCompare)
'
'   n = 0
'
'   For m = 1 + 1 To 12 + 1 ' el 1 son los nombres de los meses
'      Fact(m - 1) = 0
'      Fila = FwGetXmlTag(Page, "tr", m, t, -1)
'      If Fila <> "" Then
'         Td = FwGetXmlTag(Fila, "td", 4, 0, -1)
'         sFact = FwGetXmlTag(Td, "p", 1, 0, -1)
'         If sFact <> "" And sFact <> "&nbsp;" Then
'            Fact(m - 1) = Val(ReplaceStr(ReplaceStr(sFact, ".", ""), ",", "."))
'
'            If Abs(Fact(m - 1)) > 2 Then
'               Debug.Print "Factor inv�lido"
'               Fact(m - 1) = 1
'            End If
'
'            n = n + 1
'         End If
'      End If
'   Next m
'
'   SII_Factores_old = (n > 0)
'
'End Function

'Public Function SII_Factores(ByVal Ano As Long, Fact() As Double) As Boolean
'   Dim Path As String, Page As String, T As Long, m As Integer, n As Integer
'   Dim Fila As String, sFact As String, Td As String, m1 As Integer
'
'   SII_Factores = False
'
'   If Ano < 2000 Then
'      Exit Function
'   End If
'
'   Path = "/pagina/valores/correccion/correccion" & Ano & ".htm"
'
'   Page = FwWebReadPage("www.sii.cl", Path)
'   If Page = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
'      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
'      Exit Function
'   End If
'
'   Page = ReplaceStr(Page, "   ", " ") ' 12 ene 2015: para eliminar el problema que cambien la cantidad de blancos
'   Page = ReplaceStr(Page, "  ", " ")
'
'   T = InStr(1, Page, "<div id=""contenido""", vbTextCompare)
'   If T <= 0 Then
'      Exit Function
'   End If
'
'   T = InStr(T, Page, "<table ", vbTextCompare)
'
'   n = 0
'   m1 = 2
'
'   For m = 1 + m1 To 12 + m1 ' el 1 son los nombres de los meses
'      Fact(m - m1) = 0
'      Fila = FwGetXmlTag(Page, "tr", m, T, -1)
'      If Fila <> "" Then
'         Td = FwGetXmlTag(Fila, "td", 12, 0, -1)
'         Td = Trim(ReplaceStr(Td, "&nbsp;", ""))
'
'         If Td <> "" Then
'            Fact(m - m1) = 1 + Val(ReplaceStr(ReplaceStr(Td, ".", ""), ",", ".")) / 100
'
'            If Abs(Fact(m - m1)) > 2 Then
'               Debug.Print "Factor inv�lido"
'               Fact(m - m1) = 1
'            Else
'               n = n + 1
'            End If
'
'         End If
'      End If
'   Next m
'
'   SII_Factores = (n > 0)
'
'End Function

' 31 oct 2017: el SII cambi� la p�gina
Public Function SII_Factores(ByVal Ano As Integer, Fact() As SII_Fact_t) As Integer ' 19 ene 2017: se cambia de boolean a integer
   Dim Path As String, Page As String, T As Long, m As Integer, n As Integer, x As Integer
   Dim Fila As String, sFact As String, Td As String, m1 As Integer, Buf As String, Url As String

   SII_Factores = SERR_OK

   If Ano < 2011 Then
      SII_Factores = SERR_BADPARAM
      Exit Function
   End If
      
   Url = FwWebReadPage(URL_FWDOC & "?d=FCOR&a=" & Ano & "&u=1", "") ' u=1 para que no salte

   If Left(Url, 6) <> "##URL=" Then
      MsgBox1 "Error en la conexi�n a Internet." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline)." & vbCrLf & "Intente actualizar nuevamente.", vbExclamation
      Exit Function
   End If

   Url = Mid(Url, 7)
   
'   Path = "/pagina/renta/" & (Ano + 1) & "/grandes_contribuyentes.htm"
'   Path = "/pagina/renta/" & (Ano + 1) & "/personas_naturales.htm" ' 15 ene 2018: se cambia la p�gina
'   Path = "/valores_y_fechas/renta/" & (Ano + 1) & "/personas_naturales.html" ' 18 ene 2018: nueva p�gina
'   Page = FwWebReadPage("www.sii.cl", Path)
   
   Page = FwWebReadPage(Url, "")
   
   If Len(Page) < 20 Or modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
      SII_Factores = SERR_PGNOTFND
      Exit Function
   End If

   Page = ReplaceStr(Page, "   ", " ") ' 12 ene 2015: para eliminar el problema que cambien la cantidad de blancos
   Page = ReplaceStr(Page, "  ", " ")

   Buf = "Factores de actualizaci�n directos a�o " & Ano
   Buf = Ansi2UTF8_2(Buf)

   T = InStr(1, Page, Buf, vbTextCompare)
   If T <= 0 Then
      MsgBox1 "No existe la p�gina de 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      SII_Factores = SERR_NOINFO
      Exit Function
   End If

' en la p�gina
   T = InStr(T, Page, "<table ", vbTextCompare)
   If T <= 0 Then
      MsgBox1 "No existe la informaci�n de 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      SII_Factores = SERR_NOINFO
      Exit Function
   End If
   
   ' Buscamos donde comienzan los meses
   For x = 0 To 12
      Fila = FwGetXmlTag(Page, "tr", 1 + x, T, -1)
      If InStr(1, Fila, gNomMes(1) & " " & Ano, vbTextCompare) > 0 Then
         Exit For
      End If
   Next x

   n = 0
   m1 = 1

   For m = 1 To 12   ' el 1 son los nombres de los meses
      Fact(m).Fact = 0
      Fact(m).bFact = 0
      
      Fila = FwGetXmlTag(Page, "tr", m + x, T, -1)
      If Fila <> "" Then
         Td = FwGetXmlTag(Fila, "td", 1, 0, -1)
         
         If InStr(1, Td, gNomMes(m), vbTextCompare) > 0 Then
         
            Td = FwGetXmlTag(Fila, "td", 2, 0, -1)
            
            If Td <> "" Then
            
               Td = FwGetXmlTag(Td, "p", 1, 0, -1)
            
               Fact(m).Fact = Val(ReplaceStr(ReplaceStr(Td, ".", ""), ",", "."))
               Fact(m).bFact = 1
               
               If Fact(m).Fact < 0 Or Fact(m).Fact > 2 Then
                  Debug.Print "Factor inv�lido"
                  Fact(m).Fact = 1
               Else
                  n = n + 1
               End If
            
            End If
         End If
      End If
   Next m

   If n <= 0 Then
      SII_Factores = SERR_NODATA
      MsgBox1 "No se encontraron 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      Exit Function
   End If

   SII_Factores = SERR_OK

End Function

'Public Function SII_GetUFs(ByVal AnoMes As Long, UFs() As Double) As Boolean
'   Dim Path As String, Page As String, t As Long, d As Integer, m As Integer, n As Integer
'   Dim Fila As String, sUF As String
'
'   SII_GetUFs = False
'
'   Path = "/pagina/valores/uf/uf" & AnoMes \ 100 & ".htm"
'
'   Page = FwWebReadPage("www.sii.cl", Path)
'   If Page = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
'      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
'      Exit Function
'   End If
'
'   t = InStr(Page, "<div id=""contenido"" ")
'   If t <= 0 Then
'      Exit Function
'   End If
'
'   Page = ReplaceStr(Page, "<br />", "<br>") ' 31 mar 2014: se agrega porque se marea con /> en medio del tr
'   Page = ReplaceStr(Page, "<br/>", "<br>") ' 31 mar 2014
'
'   t = InStr(t, Page, "<table ", vbTextCompare)
'
'   m = AnoMes Mod 100
'   n = 0
'
'   For d = 1 + 1 To 31 + 1 ' el 1 son los nombres de los meses
'      UFs(d - 1) = 0
'      Fila = FwGetXmlTag(Page, "tr", d, t, -1)
'      If Fila <> "" Then
'         sUF = FwGetXmlTag(Fila, "td", m, 0, -1)
'         If sUF <> "" And sUF <> "&nbsp;" Then
'            UFs(d - 1) = Val(ReplaceStr(ReplaceStr(sUF, ".", ""), ",", "."))
'            n = n + 1
'         End If
'      End If
'   Next d
'
'   SII_GetUFs = (n > 0)
'
'End Function
'Public Function SII_Factores_old(ByVal Ano As Long, Fact() As Double) As Boolean
'   Dim Path As String, Page As String, t As Long, m As Integer, n As Integer
'   Dim Fila As String, sFact As String, Td As String
'
'   SII_Factores_old = False
'
'   If Ano < 2000 Then
'      Exit Function
'   End If
'
'   Path = "/pagina/renta/" & (Ano + 1) & "/grandes_contribuyentes.htm"
'
'   Page = FwWebReadPage("www.sii.cl", Path)
'   If Page = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
'      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
'      Exit Function
'   End If
'
'   Page = ReplaceStr(Page, "   ", " ") ' 12 ene 2015: para eliminar el problema que cambien la cantidad de blancos
'   Page = ReplaceStr(Page, "  ", " ")
'
''   t = InStr(Page, " 6.- Porcentajes y Factores de Actualizaci")
'   t = InStr(1, Page, " 6.- Porcentajes y Factores de Actualizaci", vbTextCompare) ' 12 ene 2015: porque cambiaron a mayusculas
'   If t <= 0 Then
'      Exit Function
'   End If
'
'   t = InStr(t, Page, "<table ", vbTextCompare)
'
'   n = 0
'
'   For m = 1 + 1 To 12 + 1 ' el 1 son los nombres de los meses
'      Fact(m - 1) = 0
'      Fila = FwGetXmlTag(Page, "tr", m, t, -1)
'      If Fila <> "" Then
'         Td = FwGetXmlTag(Fila, "td", 4, 0, -1)
'         sFact = FwGetXmlTag(Td, "p", 1, 0, -1)
'         If sFact <> "" And sFact <> "&nbsp;" Then
'            Fact(m - 1) = Val(ReplaceStr(ReplaceStr(sFact, ".", ""), ",", "."))
'
'            If Abs(Fact(m - 1)) > 2 Then
'               Debug.Print "Factor inv�lido"
'               Fact(m - 1) = 1
'            End If
'
'            n = n + 1
'         End If
'      End If
'   Next m
'
'   SII_Factores_old = (n > 0)
'
'End Function

'Public Function SII_Factores(ByVal Ano As Long, Fact() As Double) As Boolean
'   Dim Path As String, Page As String, T As Long, m As Integer, n As Integer
'   Dim Fila As String, sFact As String, Td As String, m1 As Integer
'
'   SII_Factores = False
'
'   If Ano < 2000 Then
'      Exit Function
'   End If
'
'   Path = "/pagina/valores/correccion/correccion" & Ano & ".htm"
'
'   Page = FwWebReadPage("www.sii.cl", Path)
'   If Page = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
'      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
'      Exit Function
'   End If
'
'   Page = ReplaceStr(Page, "   ", " ") ' 12 ene 2015: para eliminar el problema que cambien la cantidad de blancos
'   Page = ReplaceStr(Page, "  ", " ")
'
'   T = InStr(1, Page, "<div id=""contenido""", vbTextCompare)
'   If T <= 0 Then
'      Exit Function
'   End If
'
'   T = InStr(T, Page, "<table ", vbTextCompare)
'
'   n = 0
'   m1 = 2
'
'   For m = 1 + m1 To 12 + m1 ' el 1 son los nombres de los meses
'      Fact(m - m1) = 0
'      Fila = FwGetXmlTag(Page, "tr", m, T, -1)
'      If Fila <> "" Then
'         Td = FwGetXmlTag(Fila, "td", 12, 0, -1)
'         Td = Trim(ReplaceStr(Td, "&nbsp;", ""))
'
'         If Td <> "" Then
'            Fact(m - m1) = 1 + Val(ReplaceStr(ReplaceStr(Td, ".", ""), ",", ".")) / 100
'
'            If Abs(Fact(m - m1)) > 2 Then
'               Debug.Print "Factor inv�lido"
'               Fact(m - m1) = 1
'            Else
'               n = n + 1
'            End If
'
'         End If
'      End If
'   Next m
'
'   SII_Factores = (n > 0)
'
'End Function

' 13 mar 2020: Correccion monetaria del �ltimo mes con datos
Public Function SII_CorrMonet(ByVal Ano As Long, Fact() As SII_Fact_t) As Integer ' 19 ene 2017: se cambia de boolean a integer
   Dim Path As String, Page As String, T As Long, m As Integer, n As Integer, x As Integer, c As Integer
   Dim Fila As String, sFact As String, Td As String, m1 As Integer, Buf As String, Url As String, mOk As Boolean

   SII_CorrMonet = SERR_OK

   If Ano < 2013 Then
      SII_CorrMonet = SERR_BADPARAM
      Exit Function
   End If
   
   Url = FwWebReadPage(URL_FWDOC & "?d=FCORM&a=" & Ano & "&u=1", "") ' u=1 para que no salte

   If Left(Url, 6) <> "##URL=" Then
      MsgBox1 "Error en la conexi�n a Internet." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline)." & vbCrLf & "Intente actualizar nuevamente.", vbExclamation
      Exit Function
   End If

   Url = Mid(Url, 7)
      
   Page = FwWebReadPage(Url, "")
   
   If Len(Page) < 20 Or modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
      SII_CorrMonet = SERR_PGNOTFND
      Exit Function
   End If

   Page = ReplaceStr(Page, "   ", " ") ' 12 ene 2015: para eliminar el problema que cambien la cantidad de blancos
   Page = ReplaceStr(Page, "  ", " ")

   Buf = "Porcentajes de Actualizaci�n Correcci�n Monetaria (T�rmino de Giro), A�o " & Ano
   Buf = HtmlEscape2(Buf)

   T = InStr(1, Page, Buf, vbTextCompare)
   If T <= 0 Then
      MsgBox1 "No existe la p�gina de 'Correcci�n Monetaria Mensual' para el a�o " & Ano, vbExclamation
      SII_CorrMonet = SERR_NOINFO
      Exit Function
   End If

' en la p�gina
   T = InStr(T, Page, "<table class='table table-hover table-bordered'>", vbTextCompare)
   If T <= 0 Then
      MsgBox1 "No existe la informaci�n de 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      SII_CorrMonet = SERR_NOINFO
      Exit Function
   End If
   
   T = InStr(T, Page, "<tbody>", vbTextCompare)
   If T <= 0 Then
      MsgBox1 "No existe la informaci�n de 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      SII_CorrMonet = SERR_NOINFO
      Exit Function
   End If
   
   n = 0
   m1 = 1
   c = 0

   For m = 1 To 13   ' el 1 son los nombres de los meses
      Fact(m - 1).FactR = 0
      Fact(m - 1).Fact = 0
      Fact(m - 1).bFact = 0
      
      Fila = FwGetXmlTag(Page, "tr", m, T, -1)
      If Fila <> "" And (c = 0 Or m - 1 <= c) Then
      
         If m = 1 Then ' buscamos �ltimo mes
            For c = 1 To 12
               Td = Trim(FwGetXmlTag(Fila, "td", c, 0, -1))
               If Td = "" Then
                  Exit For
               End If
            Next c
            c = c - 1
         End If
            
         If m > 1 Then
            Td = FwGetXmlTag(Fila, "th", 1, 0, -1)
            mOk = (InStr(1, Td, gNomMes(m - 1), vbTextCompare) > 0)
         Else
            mOk = True
         End If
      
         If mOk Then
            Td = Trim(FwGetXmlTag(Fila, "td", c, 0, -1))
            If Td <> "" Then
               Fact(m - 1).FactR = Val(ReplaceStr(Td, ",", "."))
               Fact(m - 1).FactR = 1 + Fact(m - 1).FactR / 100
               Fact(m - 1).Fact = IIf(Fact(m - 1).FactR < 1, 1, Fact(m - 1).FactR) ' 2 abr 2020: seg�n Victor Morales, si es negativo no hay reajuste
               
               Fact(m - 1).bFact = 1
               n = n + 1
            End If
         End If

      End If
   Next m

   If n <= 0 Then
      SII_CorrMonet = SERR_NODATA
      MsgBox1 "No se encontraron 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      Exit Function
   End If

   SII_CorrMonet = SERR_OK

End Function

' 13 mar 2020: Correccion monetaria del a�o (mat por Matriz)  (i, j): i: fila, j: columna
Public Function SII_CorrMonetAnual(ByVal Ano As Long, Fact() As SII_Fact_t) As Integer
   Dim Path As String, Page As String, T As Long, m As Integer, n As Integer, x As Integer, c As Integer, i As Integer, j As Integer
   Dim Fila As String, sFact As String, Td As String, m1 As Integer, Buf As String, Url As String, mOk As Boolean

   SII_CorrMonetAnual = SERR_OK

   If Ano < 2013 Then
      SII_CorrMonetAnual = SERR_BADPARAM
      Exit Function
   End If
   
   Url = FwWebReadPage(URL_FWDOC & "?d=FCORM&a=" & Ano & "&u=1", "") ' u=1 para que no salte

   If Left(Url, 6) <> "##URL=" Then
      MsgBox1 "Error en la conexi�n a Internet." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline)." & vbCrLf & "Intente actualizar nuevamente.", vbExclamation
      Exit Function
   End If

   Url = Mid(Url, 7)
      
   Page = FwWebReadPage(Url, "")
   
   If Len(Page) < 20 Or modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
      SII_CorrMonetAnual = SERR_PGNOTFND
      Exit Function
   End If

   Page = ReplaceStr(Page, "   ", " ") ' 12 ene 2015: para eliminar el problema que cambien la cantidad de blancos
   Page = ReplaceStr(Page, "  ", " ")

   Buf = "Porcentajes de Actualizaci�n Correcci�n Monetaria (T�rmino de Giro), A�o " & Ano
   Buf = HtmlEscape2(Buf)

   T = InStr(1, Page, Buf, vbTextCompare)
   If T <= 0 Then
      MsgBox1 "No existe la p�gina de 'Correcci�n Monetaria Mensual' para el a�o " & Ano, vbExclamation
      SII_CorrMonetAnual = SERR_NOINFO
      Exit Function
   End If

' en la p�gina
   T = InStr(T, Page, "<table class='table table-hover table-bordered'>", vbTextCompare)
   If T <= 0 Then
      MsgBox1 "No existe la informaci�n de 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      SII_CorrMonetAnual = SERR_NOINFO
      Exit Function
   End If
   
   T = InStr(T, Page, "<tbody>", vbTextCompare)
   If T <= 0 Then
      MsgBox1 "No existe la informaci�n de 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      SII_CorrMonetAnual = SERR_NOINFO
      Exit Function
   End If
   
   n = 0
   m1 = 1
   c = 0

   For i = 0 To 12
      For j = 0 To 12
         Fact(i, j).Fact = 0
         Fact(i, j).bFact = 0
      Next j
   Next i
         
   For m = 1 To 13   ' el 0 son los nombres de los meses de la fila superior
      
      If c > 0 And m - 1 > c Then
         Exit For
      End If

      Fila = FwGetXmlTag(Page, "tr", m, T, -1)
      If Fila <> "" And (c = 0 Or m - 1 <= c) Then
      
         If m = 1 Then ' buscamos �ltimo mes
            For c = 1 To 12
               Td = Trim(FwGetXmlTag(Fila, "td", c, 0, -1))
               If Td = "" Then
                  Exit For
               End If
            Next c
            c = c - 1
         End If
                        
         If m > 1 Then
            Td = FwGetXmlTag(Fila, "th", 1, 0, -1)
            mOk = (InStr(1, Td, gNomMes(m - 1), vbTextCompare) > 0)
         Else
            mOk = True
         End If
      
         If mOk Then
            For j = m - 1 To c
               Td = Trim(FwGetXmlTag(Fila, "td", j, 0, -1))
               If Td <> "" Then
                  Fact(m - 1, j).FactR = Val(ReplaceStr(Td, ",", "."))
                  Fact(m - 1, j).FactR = 1 + Fact(m - 1, j).FactR / 100
                  Fact(m - 1, j).Fact = IIf(Fact(m - 1, j).FactR < 1, 1, Fact(m - 1, j).FactR) ' 2 abr 2020: seg�n Victor Morales, si es negativo no hay reajuste

                  Fact(m - 1, j).bFact = 1
                  n = n + 1
               End If
            Next j
         End If

      End If
   Next m

   If n <= 0 Then
      SII_CorrMonetAnual = SERR_NODATA
      MsgBox1 "No se encontraron 'Factores de Actualizaci�n Directos' para el a�o " & Ano, vbExclamation
      Exit Function
   End If

   SII_CorrMonetAnual = SERR_OK

End Function

'Public Sub Test1(ByVal Ano As Integer)
'   Dim Fact(12, 12) As SII_Fact_t, i As Integer, j As Integer, Buf As String
'
'   Call SII_CorrMonetAnual(Ano, Fact)
'
'   For i = 0 To 12
'      Buf = ""
'      For j = 1 To 12
'         Buf = Buf & vbTab & IIf(Fact(i, j).bFact, Format(Fact(i, j).Fact, "0.000"), "     ")
'      Next j
'
'      Debug.Print i & ") " & Buf
'   Next i
'
'End Sub


' IPCs( m, c ) es una matriz, m es el mes
' If bUTM = True
'  c = 0: UTM
'  c = 1: si tiene UTM
' Else
'  c = 0: var % de IPC
'  c = 1: si tiene var % de IPC
'  c = 2: puntos de IPC
'  c = 3: si tiene puntos de IPC
'  c = 4: var % de IPC acumulado
'  c = 5: si tiene var % de IPC acumulado
' 12 sep 2017: el SII cambia la p�gina y el formato de la p�gina
Public Function SII_GetIPCs(ByVal Ano As Integer, IPCs() As SII_IPC_t, Optional ByVal bUTM As Boolean = False) As Integer ' 19 ene 2018: cambia de boolean a integer
   Dim Path As String, Page As String, T As Long, d As Integer, m As Integer, n As Integer
   Dim Fila As String, sIPC As String, Url As String

   SII_GetIPCs = SERR_OK
   'URL_FWDOC = Replace(URL_FWDOC, "http", "https")
   Url = FwWebReadPage(URL_FWDOC & "?d=UTM&a=" & Ano & "&u=1", "") ' u=1 para que no salte
    Url = Replace(Url, "http", "https")
   If Left(Url, 6) <> "##URL=" Then
      MsgBox1 "Error en la conexi�n a Internet." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline)." & vbCrLf & "Intente actualizar nuevamente.", vbExclamation
      Exit Function
   End If

   Url = Mid(Url, 7)

'   Path = "/valores_y_fechas/utm/utm" & Ano & ".htm"
'   Page = FwWebReadPage("www.sii.cl", Path)

'   Dim myMSXML As Object
'   Set myMSXML = CreateObject("Microsoft.XmlHttp")
'   myMSXML.Open "POST", Url, False
''myMSXML.setRequestHeader "x-api-key", ApiKey
'myMSXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'myMSXML.send
'MsgBox myMSXML.responseText

'Dim WEB As WebBrowser
'DIM HTTP AS
'
'Set WEB = New WebBrowser ' Se instancia el WebBrower
'WEB.ScriptErrorsSuppressed = True ' Oculta la ventana de errores si alg�n script de la p�gina fall� (de todas formas no los necesitamos)
'        WEB.Navigate2 (Url) ' Carga la p�gina web creando un nuevo documento HTML
        
        Dim MyRequest As Object
        'Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        'Set MyRequest = CreateObject("Msxml2.ServerXMLHTTP.6.0")
        Set MyRequest = CreateObject("MSXML2.XMLHTTP")
    MyRequest.Open "GET", Url, False
    'MyRequest.Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_TLS1_1
    'MyRequest.Option(9) = 2056
    ' Send Request.
    MyRequest.Send

    'And we get this response
    Page = MyRequest.ResponseText

   
   'Page = FwWebReadPage(Url, "")
   If Len(Page) < 20 Or modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
      SII_GetIPCs = SERR_PGNOTFND
      Exit Function
   End If
   
   T = InStr(Page, "En la siguiente tabla se presenta para los meses del " & Ano)
   If T <= 0 Then
      SII_GetIPCs = SERR_NOINFO
      MsgBox1 "No existe la p�gina con los datos de IPCs del a�o " & Ano, vbExclamation
      Exit Function
   End If
   
   Page = ReplaceStr(Page, "<br />", "<br>") ' 31 mar 2014: se agrega porque se marea con /> en medio del tr
   Page = ReplaceStr(Page, "<br/>", "<br>") ' 31 mar 2014
   
   T = InStr(T, Page, "<table ", vbTextCompare)
   n = 0

   For d = 1 + 2 To 12 + 2 ' el 1 son los nombres de los meses y el 2 detalle
      IPCs(d - 2).UTM = 0
      IPCs(d - 2).bUTM = 0   ' no tiene valor
      
      IPCs(d - 2).VarIpc = 0
      IPCs(d - 2).bVarIpc = 0   ' no tiene valor
      
      If bUTM = False Then
         IPCs(d - 2).PIpc = 0
         IPCs(d - 2).bPIpc = 0   ' no tiene valor
         
         IPCs(d - 2).bVarAcum = 0
         IPCs(d - 2).bVarAcum = 0   ' no tiene valor
      End If
      
      Fila = FwGetXmlTag(Page, "tr", d, T, -1)
      If Fila <> "" Then
      
         If InStr(1, Fila, gNomMes(d - 2), vbTextCompare) > 0 Then
      
            If bUTM Then
               sIPC = FwGetXmlTag(Fila, "td", 1, 0, -1)  ' UTM
               If sIPC <> "" And sIPC <> "&nbsp;" Then
                  IPCs(d - 2).UTM = Val(ReplaceStr(sIPC, ".", ""))
                  IPCs(d - 2).bUTM = 1   ' tiene valor
                  n = n + 1
               End If
            Else
               sIPC = FwGetXmlTag(Fila, "td", 4, 0, -1)  ' % de IPC
               If sIPC <> "" And sIPC <> "&nbsp;" Then
                  IPCs(d - 2).VarIpc = Val(ReplaceStr(sIPC, ",", "."))  ' 0: Variaci�n Mensual
                  IPCs(d - 2).bVarIpc = 1   ' tiene valor
                  n = n + 1
               End If
               
               sIPC = FwGetXmlTag(Fila, "td", 3, 0, -1)  ' Puntos de IPC
               If sIPC <> "" And sIPC <> "&nbsp;" Then
                  IPCs(d - 2).PIpc = Val(ReplaceStr(ReplaceStr(sIPC, ".", ""), ",", ".")) ' 2: Puntos de IPC
                  IPCs(d - 2).bPIpc = 1   ' tiene valor
               End If
               
               sIPC = FwGetXmlTag(Fila, "td", 5, 0, -1)  ' % de IPC Acumulado
               If sIPC <> "" And sIPC <> "&nbsp;" Then
                  IPCs(d - 2).VarAcum = Val(ReplaceStr(sIPC, ",", ".")) ' 4: 0: Variaci�n Mensual Acumulada
                  IPCs(d - 2).bVarAcum = 1   ' tiene valor
               End If
            End If
         End If
      End If
   Next d

   If n <= 0 Then
      SII_GetIPCs = SERR_NODATA
      MsgBox1 "No se encontraron datos de IPCs o UTMs para el a�o " & Ano, vbExclamation
      Exit Function
   End If
   
   SII_GetIPCs = SERR_OK
   
End Function

'Public Function SII_GetIPCs(ByVal Ano As Integer, IPCs() As Double, Optional ByVal bUTM As Boolean = False) As Boolean
'   Dim Path As String, Page As String, t As Long, d As Integer, m As Integer, n As Integer
'   Dim Fila As String, sIPC As String
'
'   SII_GetIPCs = False
'
'   Path = "/pagina/valores/utm/utm" & Ano & ".htm"
'
'   Page = FwWebReadPage("www.sii.cl", Path)
'   If Page = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
'      MsgBox1 "Error en la conexi�n." & vbCrLf & "Verifique en Internet Explorer que no est� trabajando desconectado (Offline).", vbExclamation
'      Exit Function
'   End If
'
'   t = InStr(Page, "<div id=""contenido"">")
'   If t <= 0 Then
'      Exit Function
'   End If
'
'   Page = ReplaceStr(Page, "<br />", "<br>") ' 31 mar 2014: se agrega porque se marea con /> en medio del tr
'   Page = ReplaceStr(Page, "<br/>", "<br>") ' 31 mar 2014
'
'   t = InStr(t, Page, "<table ", vbTextCompare)
'   n = 0
'
'   For d = 1 + 2 To 12 + 2 ' el 1 son los nombres de los meses y el 2 detalle
'      IPCs(d - 2, 0) = 0
'      IPCs(d - 2, 1) = 0   ' no tiene valor
'
'      If bUTM = False Then
'         IPCs(d - 2, 2) = 0
'         IPCs(d - 2, 3) = 0   ' no tiene valor
'      End If
'
'      Fila = FwGetXmlTag(Page, "tr", d, t, -1)
'      If Fila <> "" Then
'
'         If bUTM Then
'            sIPC = FwGetXmlTag(Fila, "td", 1, 0, -1)  ' UTM
'            If sIPC <> "" And sIPC <> "&nbsp;" Then
'               IPCs(d - 2, 0) = Val(ReplaceStr(sIPC, ".", ""))
'               IPCs(d - 2, 1) = 1   ' tiene valor
'               n = n + 1
'            End If
'         Else
'            sIPC = FwGetXmlTag(Fila, "td", 4, 0, -1)  ' % de IPC
'            If sIPC <> "" And sIPC <> "&nbsp;" Then
'               IPCs(d - 2, 0) = Val(ReplaceStr(sIPC, ",", "."))
'               IPCs(d - 2, 1) = 1   ' tiene valor
'               n = n + 1
'            End If
'
'            sIPC = FwGetXmlTag(Fila, "td", 3, 0, -1)  ' Puntos de IPC
'            If sIPC <> "" And sIPC <> "&nbsp;" Then
'               IPCs(d - 2, 2) = Val(ReplaceStr(ReplaceStr(sIPC, ".", ""), ",", "."))
'               IPCs(d - 2, 3) = 1   ' tiene valor
'            End If
'         End If
'      End If
'   Next d
'
'   SII_GetIPCs = (n > 0)
'
'End Function

Public Function EmpresaSii(Rut As String, Clave As String) As EmpresaSII_t
'********** INICIO ****************
   Dim Params As String
   Dim Url As String
   Dim Resp As String
   Dim Termina As Boolean
   Dim i As Integer
   Dim x As Integer
   Dim v As Integer
   Dim Info() As String
   Dim Traza As String
   
   Dim RespAux As String
   Dim CodError As Integer
   Dim DesError As String
   Dim InfoEmpresa As EmpresaSII_t

'   Dim Nombre, ApellidoPaterno, ApellidoMaterno, RazonSocial As String
'   Dim Calle, Numero, Dpto, villaPoblacion As String
'   Dim Region, Comuna, Ciudad As String
'   Dim AreaTelefono, Telefono, telefonoMovil, AreaFax, Fax, email As String
'   Dim fechaConstitucion, fechaInicioActividades As Long
'   Dim glosaActividadGiro As String
'   Dim CodActEconom As Long
'   Dim RutRepresentante, NombreRepresentante, Vigente As String
'   Dim Representante(1) As Representante_t
   Dim ContaRepres As Integer

   v = 0

    If W.InDesign Then
        'Rut = "77765060-2"
        'Clave = "romal1"
        'lAno = "2023"
        'lMes = "04"
    End If
'
   'Sleep (10000)
   'Params = "rut=17533256&dv=1&referencia=https%3A%2F%2Fmisiir.sii.cl%2Fcgi_misii%2Fsiihome.cgi&411=%20&rutcntr=17533256-1&clave=Fpriet4512"
   Params = "rut=" & vFmtCID(Rut) & "&dv=" & DV_Rut(vFmtCID(Rut)) & "&referencia=https%3A%2F%2Fmisiir.sii.cl%2Fcgi_misii%2Fsiihome.cgi&411=%20&rutcntr=" & vFmtCID(Rut) & "-" & DV_Rut(vFmtCID(Rut)) & "&clave=" & Clave
   Url = URL_SII_LOGIN
   Resp = FwPostPageSII2(Url, Params, "application/x-www-form-urlencoded", SII_LOGIN)
   'La_Title = gLexContab
   If Val(InStr(1, Resp, "titulo", vbTextCompare)) > 0 Then
        'MsgBox1 Replace(Utf8Ansi(Trim(ReplaceStr(ReplaceStr(ReplaceStr(ReplaceStr(GetMensajeSII(FwGetXmlTag(Resp, "div", 1)), "<br>", ""), "<p>", ""), vbLf, ""), vbCr, ""))), "                ", vbLf), vbInformation
        MsgBox1 "Error con la informacion ingresada, Favor verificar su Clave"
        Exit Function
   End If
   
   Params = "opc=112"
   Url = "https://misiir.sii.cl/cgi_misii/CViewCarta.cgi"
   Resp = FwPostPageSII2(Url, Params, "application/x-www-form-urlencoded", SII_RETENCIONES)
   
  Call recorreJson(Resp, "codigoError", 1)
  
  If recorreJson(Resp, "codigoError", 1) = "0" Then
    
    InfoEmpresa.CodError = recorreJson(Resp, "codigoError", 1)
    If recorreJson(Resp, "nombres", 1) = "" Then
        InfoEmpresa.RazonSocial = recorreJson(Resp, "razonSocial", 1)
    Else
        InfoEmpresa.Nombre = recorreJson(Resp, "nombres", 1)
        InfoEmpresa.ApellidoPaterno = recorreJson(Resp, "apellidoPaterno", 1)
        InfoEmpresa.ApellidoMaterno = recorreJson(Resp, "apellidoMaterno", 1)
    End If
    
    InfoEmpresa.telefonoMovil = recorreJson(Resp, "telefonoMovil", 1)
    InfoEmpresa.email = recorreJson(Resp, "eMail", 1)
    InfoEmpresa.glosaActividadGiro = recorreJson(Resp, "glosaActividad", 1)
    
    If recorreJson(Resp, "fechaConstitucion", 1) <> "" Then
        InfoEmpresa.fechaConstitucion = CLng(CDate(Mid(recorreJson(Resp, "fechaConstitucion", 1), 1, 10)))
    End If
    If recorreJson(Resp, "fechaInicioActividades", 1) <> "" Then
        InfoEmpresa.fechaInicioActividades = CLng(CDate(Mid(recorreJson(Resp, "fechaInicioActividades", 1), 1, 10)))
    End If
    'fechaInicioActividades = recorreJson(Resp, "fechaInicioActividades", 1)
    
    RespAux = Mid(Resp, Val(InStr(Val(InStr(1, Resp, "direcciones", vbTextCompare)), Resp, "[", vbTextCompare)), Val(InStr(Val(InStr(1, Resp, "direcciones", vbTextCompare)), Resp, "]", vbTextCompare)) - Val(InStr(Val(InStr(1, Resp, "direcciones", vbTextCompare)), Resp, "[", vbTextCompare)))
    
    For i = 1 To CuentaCaracteres(RespAux, "}")
        If recorreJson(RespAux, "codigoError", i) = "0" Then
            
            If recorreJson(RespAux, "tipoDomicilioCodigo", i) = "1" Then
                InfoEmpresa.Calle = recorreJson(RespAux, "calle", i)
                InfoEmpresa.Numero = recorreJson(RespAux, "numero", i)
                InfoEmpresa.Dpto = recorreJson(RespAux, "departamento", i)
                InfoEmpresa.villaPoblacion = recorreJson(RespAux, "villaPoblacion", i)
                InfoEmpresa.Region = recorreJson(RespAux, "regionCodigo", i)
                InfoEmpresa.Comuna = recorreJson(RespAux, "ciudad", i)
                InfoEmpresa.Ciudad = recorreJson(RespAux, "ciudad", i)
                InfoEmpresa.Fax = IIf(recorreJson(RespAux, "areaFax", i) = "", "", recorreJson(RespAux, "areaFax", i)) & IIf(recorreJson(RespAux, "fax", i) = "", "", recorreJson(RespAux, "fax", i))
                InfoEmpresa.Telefono = IIf(recorreJson(RespAux, "areaTelefono", i) = "", "", recorreJson(RespAux, "areaTelefono", i)) & IIf(recorreJson(RespAux, "telefono", i) = "", "", recorreJson(RespAux, "telefono", i))
                Exit For
            End If
            
        End If
    Next i
    
    
    RespAux = Mid(Resp, Val(InStr(Val(InStr(1, Resp, "actEcos", vbTextCompare)), Resp, "[", vbTextCompare)), Val(InStr(Val(InStr(1, Resp, "actEcos", vbTextCompare)), Resp, "]", vbTextCompare)) - Val(InStr(Val(InStr(1, Resp, "actEcos", vbTextCompare)), Resp, "[", vbTextCompare)))
    
    For i = 1 To CuentaCaracteres(RespAux, "}")
        If recorreJson(RespAux, "codigoError", i) = "0" Then
            
            'If recorreJson(RespAux, "tipoDomicilioCodigo", i) = "1" Then
                InfoEmpresa.CodActEconom = IIf(recorreJson(RespAux, "codigo", i + 1) = "", 0, CLng(recorreJson(RespAux, "codigo", i + 1)))
                Exit For
            'End If
            
        End If
    Next i
    
    RespAux = Mid(Resp, Val(InStr(Val(InStr(1, Resp, "representantes", vbTextCompare)), Resp, "[", vbTextCompare)), Val(InStr(Val(InStr(1, Resp, "representantes", vbTextCompare)), Resp, "]", vbTextCompare)) - Val(InStr(Val(InStr(1, Resp, "representantes", vbTextCompare)), Resp, "[", vbTextCompare)))
    
    ContaRepres = 0
    For i = 1 To CuentaCaracteres(RespAux, "}")
        If recorreJson(RespAux, "codigoError", i) = "0" Then
            
            If Mid(recorreJson(RespAux, "vigente", i), 1, 1) = "S" Then
                InfoEmpresa.Representante(ContaRepres).Rut = IIf(recorreJson(RespAux, "rut", i) = "", 0, CLng(recorreJson(RespAux, "rut", i)))
                
                If recorreJson(RespAux, "razonSocial", i) = "" Then
                    InfoEmpresa.Representante(ContaRepres).Nombre = IIf(recorreJson(RespAux, "nombres", i) = "", "", recorreJson(RespAux, "nombres", i))
                    InfoEmpresa.Representante(ContaRepres).Nombre = InfoEmpresa.Representante(ContaRepres).Nombre & " " & IIf(recorreJson(RespAux, "apellidoPaterno", i) = "", "", recorreJson(RespAux, "apellidoPaterno", i))
                    InfoEmpresa.Representante(ContaRepres).Nombre = InfoEmpresa.Representante(ContaRepres).Nombre & " " & IIf(recorreJson(RespAux, "apellidoMaterno", i) = "", "", recorreJson(RespAux, "apellidoMaterno", i))
                Else
                    InfoEmpresa.Representante(ContaRepres).Nombre = IIf(recorreJson(RespAux, "razonSocial", i) = "", 0, CLng(recorreJson(RespAux, "razonSocial", i)))
                End If
                
                ContaRepres = ContaRepres + 1
                If ContaRepres = 2 Then
                    Exit For
                End If
            End If
            
        End If
    Next i
    
    RespAux = Mid(Resp, Val(InStr(Val(InStr(1, Resp, "socios", vbTextCompare)), Resp, "[", vbTextCompare)), Val(InStr(Val(InStr(1, Resp, "socios", vbTextCompare)), Resp, "]", vbTextCompare)) - Val(InStr(Val(InStr(1, Resp, "socios", vbTextCompare)), Resp, "[", vbTextCompare)))
    
    Dim Cant As Integer
    Cant = CuentaCaracteres(RespAux, "}") - 1
    'Dim Socios(10) As Representante_t
    ContaRepres = 0
    For i = 1 To CuentaCaracteres(RespAux, "}")
        If recorreJson(RespAux, "codigoError", i) = "0" Then
            
            If Mid(recorreJson(RespAux, "vigente", i), 1, 1) = "S" Then
                InfoEmpresa.Socios(ContaRepres).Rut = IIf(recorreJson(RespAux, "rut", i) = "", 0, CLng(recorreJson(RespAux, "rut", i)))
                
                If recorreJson(RespAux, "razonSocial", i) = "" Then
                    InfoEmpresa.Socios(ContaRepres).Nombre = IIf(recorreJson(RespAux, "nombres", i) = "", "", recorreJson(RespAux, "nombres", i))
                    InfoEmpresa.Socios(ContaRepres).Nombre = InfoEmpresa.Socios(ContaRepres).Nombre & " " & IIf(recorreJson(RespAux, "apellidoPaterno", i) = "", "", recorreJson(RespAux, "apellidoPaterno", i))
                    InfoEmpresa.Socios(ContaRepres).Nombre = InfoEmpresa.Socios(ContaRepres).Nombre & " " & IIf(recorreJson(RespAux, "apellidoMaterno", i) = "", "", recorreJson(RespAux, "apellidoMaterno", i))
                Else
                    InfoEmpresa.Socios(ContaRepres).Nombre = IIf(recorreJson(RespAux, "razonSocial", i) = "", "", recorreJson(RespAux, "razonSocial", i))
                End If
                
                InfoEmpresa.Socios(ContaRepres).porcentaje = IIf(recorreJson(RespAux, "participacionUtilidades", i) = "", 0, CLng(recorreJson(RespAux, "participacionUtilidades", i)))
                
                
                ContaRepres = ContaRepres + 1
            End If
            
        End If
    Next i
    
  End If
  
  Url = "https://misiir.sii.cl/cgi_misii/siihome.cgi"
   Resp = FwPostPageSII2(Url, "", "application/x-www-form-urlencoded", SII_RETENCIONES)
  RespAux = Mid(Resp, Val(InStr(1, Resp, "DatosCntrNow", vbTextCompare)))
  If recorreJson(RespAux, "codigoError", 1) = "0" Then
    RespAux = Mid(RespAux, Val(InStr(Val(InStr(1, RespAux, "atributos", vbTextCompare)), RespAux, "[", vbTextCompare)), Val(InStr(Val(InStr(1, RespAux, "atributos", vbTextCompare)), RespAux, "]", vbTextCompare)) - Val(InStr(Val(InStr(1, RespAux, "atributos", vbTextCompare)), RespAux, "[", vbTextCompare)))
    For i = 1 To CuentaCaracteres(RespAux, "}")
       InfoEmpresa.TipoFranqui = ValBuscaFranquicie(recorreJson(RespAux, "descAtrCodigo", i))
       If InfoEmpresa.TipoFranqui <> "0" Then
         Exit For
       End If
    Next i
  End If
  
  
  
  EmpresaSii = InfoEmpresa

   'TRAZA =
   
   Url = URL_SII_LOGOUT
   Resp = FwPostPageSII2(Url, "", "application/x-www-form-urlencoded", SII_LOGOUT)
   'La_Title = gLexContab
   If Val(InStr(1, Resp, "titulo", vbTextCompare)) > 0 Then
        MsgBox1 GetMensajeSII(FwGetXmlTag(Resp, "div", 1)), vbInformation
        'Exit Sub
   End If
   
   
'********* FIN ***********
End Function


