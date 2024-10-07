Attribute VB_Name = "Actualizaciones"
Option Explicit

'� �
'11/04/2017 Se agrega opci�n para ingresar manualmente otros ingresos y egresos a Libro de Caja
'           Se solucionan errores en los filtros del Listar Libro de Caja
'           Se corrigen errores con el c�lculo de IVA Irrecuperable en Resumen de Libros Auxiliares, Resumen de IVA y Listar Libro de Compras
'           Se agrega columna IVA Irrecuperable a Libro de Caja
'           Se modifica importaci�n de Libro de Compras en la ventana de Libro de Caja, para que incluya el IVA Irrecuperable
'12/04/2017 Se modifica validaciones del libro de caja, especialmente para RUT y Glosa (tiene que ser de al menos largo 1)
'           Se agrega opciones de vista/edici�n al Libro de Caja
'13/04/2017 Se agrega indicador de entidad relacionada en ficha entidad y selecci�n de campo ent. relacionada al ingresar RUT o importar Doc en Libro de Caja
'20/04/2017 Se libera versi�n 6.0.2
'24/04/2017 Se agrega una validaci�n cuando se selecciona 14 TER en ventana Empresa, dependiendo de tipo contribuyente
'           Se modifica la edici�n de una cuenta contable para:
'              - Agregar atributo Ajuste 14 TER
'              - Agregar c�digo asociado a cuenta contable con atributo 14 TER
'27/04/2017 Se agrega lista de Socios, Propietarios y Comuneros a ventana de edici�n de Empresa
'05/05/2017 Se agrega exportaci�n de libro de compras para Facturaci�n Electr�nica (Acepta)
'10/05/2017 Se libera versi�n 6.0.2 de nuevo....
'11-05-2017 Se corrige error de password en importaci�n de Libro de Compras o Ventas desde Facturaci�n
'16-05-2017 Se agrega posibilidad de agregar documentos de otros meses al libro de caja. Para esto se maneja Fecha de Operaci�n y Fecha Ingreso al Libro (oculta)
'17-05-2017 Se agrega opci�n de RUT = VARIOS para Libro de Caja Egresos/Ingresos para otros Ingresos o Egresos
'           Se agrega validaci�n para Libro de Compras - 'OTC' para RUT Obligatorio
'           Se modifica Importaci�n de documentos a Libro de Caja para que a IVA se le reste IVA Irrecuperable, si no quedaban duplicados
'           Se agrega NumDoc = "NumDoc - NumDocHasta" en el caso que haya NumDocHats. Eso para declarar m�s de una boleta
'           Se formatea RUT en Libro de Caja para el caso de boletas o docs que no exigen RUT
'           Se corrige error en importaci�n desde facturaci�n que hac�a calzar el ID de la Entidad en vez del Rut de la entidad al actualizar las entidades
'17/06/2017 Se modifica formato Tasa en la generaci�n del Libro Electr�nico de Compras. El entregado anteriormente en especificaciones no correspond�a
'           Se agrega validaci�n al importar Libro de Ventas, para que no acepte BOV con exento
'           Se agrega N�Fiscal Impresora, N� Informe Z, Ventas Acum Informe Z y Cantidad de Boletas a la importaci�n del libro de Ventas
'           Se agrega Cod. Cuenta Otros Imp a la importaci�n del libro de Compras y Ventas
'           Se agrega identificaci�n de centro de gesti�n cuando se importa desde Remuneraciones, para el caso de seleccionar un centro de gesti�n.
'           Se agrega impedimento para ingresar Afecto e IVA en documento M�quina Registradora.
'09/06/2017 Se correige error de Inf. Anal�tico para que oculte cuenta completa si Saldo de la cuenta es cero, cuando se selecciona la opci�n Saldos Vigentes
'           Se agrega Columna Cod. Cuenta Activo Fijo a importaci�n de Activos Fijos
'16/06/2017 Se agrega identificaci�n de centros de costo cuando se importa desde Remuneraciones, tambi�n para el caso de no seleccionar ning�n centro de costo
'           Se agrega manual de configuraci�n y manual de importaci�n de IEC para que se abra desde las ventanas respectivas
'20/06/2017 Se agregan campos a TipoDocs para controlar y validar captura de documentos en libro de ventas
'28/06/2017 Se agrega obligatoriedad de fecha de vencimiento en libro de compras y ventas.
'              Se agrega fecha de vencimiento en Libro de Retenciones y se hace obligatorio
'              Se agrega parche para asignar fecha de vencimiento a 30 d�as a los documentos que tienen la fecha de vencimiento en cero.
'              Esto se hace al momento de desplegar el libro. Al grabar el libro queda guardada la fecha de vencimiento asignada.
'              Para los documentos pendientes el usuario le asigna la fecha al editar el libro, dado que no permite grabar si no est� la fecha de vencimiento ingresada.
'              Cuando se agrega un nuevo documento, el sistema asigna por omisi�n vencimiento a 30 d�as al momento de ingresar la fecha de emisi�n
'           Se agrega fecha de vencimiento a la importaci�n del libro de retenciones
'04/07/2017 Libro de Compras y Ventas: fecha de vencimiento = fecha de emisi�n si es Entidad Relacionada, si no es pago en cuotas (s�lo para 14 TER)
'              Se asigna fecha de vencimiento = fecha de emsi�n si es Entidad Relacionada en: sel entidad, nueva entidad, ingreso de RUT, ingreso o cambiuo de fecha de emisi�n
'              Libro de Caja ingreso manual de ventas: si es Ent Relacionada se sugiere fecha de exigibilidad de pago contado, si no, a 30 d�as
'              Se agrega EntRelacionada a Documento y se copia en corrige base Documento.EntRelacionada = Entidades.EntRelacionada
'           En el caso de los ingresos (ventas), si la entidad es Relacionada, el monto del documento se da de inmediato como percibido, de acuerdo a la ley
'11/07/2017 CAMBIO TOTAL: Fecha Vencimiento:
'              Cuando se agrega documento: si es 14TER se asigna fecha vencimiento contado (fecha Emisi�n) y si no, pago a 30 d�as
'              Si es pago en cuotas, se usa la fecha de vencimiento de la primera cuota como fecha de vencimiento del doc
'              Ya no se debe dar como percibido monto de venta a ent. relacionada. El percibido es a medida que paguen, igual que en los no de ent. relacionada.
'12/07/2017 Se corrige error de ingreso de cuenta afecta y exenta para Boletas de venta en libro de CompraVentas
'18/07/2017 Se corrige error al guardar docs en libro de compras-ventas. Ent. Relacionada no se actualizaba correctamente
'           Se permite editar Cod. F22 en listado de cuentas
'           Se modifica fomrato de Fecha Oper y Fecha Exig. de Pago en libro de Caja. Debe ser dd/mm/yyyy
'           Se cambia "Ent. Relacionada" por "Norma Relaci�n 14 TER"
'           Se agrega mensaje de advertencia al modificar la calificaci�n de entidad relacionada en la ventana de edici�n de Entidad
'24/07/2017 Se agrega ventana de ingreso de cuotas. Tiene bot�n para generar cuotas en forma autom�tica
'           Se agrega pago de cuotas autom�tico, asociado a la generaci�n de pago de documentos.
'           Se agrega reporte de cuotas pagadas o impagas.
'01/08/2017 Se modifica importaci�n de compras y ventas en Libro de Caja para que transforme INGRESOS a EGRESOS y viceversa
'           Se modifica Importaci�n de Retenciones a Libro de Caja para adecuarla a nuevos requerimientos
'08/08/2017 Se agrega funci�n de importaci�n al Libro de Caja de Otros Ingresos y Egresos desde comprobantes, incluyendo llenado de columna Peribido o Pagado
'           Se elimina de Libro de Egresos la restricci�n de no incorporar documentos de Retenci�n. Ahora si se incluyen estos documentos.
'10/08/2017 Se corrige error en importaci�n de documentos a Libro de Caja que se produc�a al restar el IVA Irreecuperabre al IVA cuando el IVA Irrecuperable era NULO
'           Se modifica importaci�n de documentos a Libro de Caja para que no pise la Descripci�n del Documento en el Libro de Caja, si �sta es distinto de blanco. Mantiene la que ingres� el usuario en el libro de caja.
'           Se agrega NCF a la importaci�n de Compras en Libro de Caja para que la pase a Ingreso
'           Se bloquea ingreso de cuotas a notas de cr�dito y d�bito
'11/08/2017 Se bloquea el pago de cuotas saltadas o que no sea la primera impaga
'           Se corrige funci�n de generaci�n de cuotas por error al generar m�s de 28 cuotas.
'16/08/2017 Se corrige error de formateo de RUT en Libro de Caja para el caso de documentos de exportaci�n, en que el RUT es el de la empresa misma
'17/08/2017 Se agrega Descripci�n obligatoria en libros de Compras, Ventas y Retenciones si el empresa 14TER, para que luego se lleve a Libro de Caja donde la descripci�n es obligatoria
'           Se permite editar s�lo la columna Operaci�n Devengada del Libro de Caja, para documentos importados de libros.
'31/08/2017 Se agrega el tema de los ingresos percibidos y egresos pagados al libro de caja
'           Se agregan mensajes por el tema del libro electr�nico de compras que ya no se utilizar� a partir de febrero 2018
'           Se agrega un campo a la edici�n de documentos para indicar si es una compra de un bien raiz
'           Se agrega DVB a la importaci�n de Ventas en Libro de Caja para que la pase a Egreso
'05/09/2017 Se agrega campos Ingreso y Egreso a tabla LibroCaja para almacenar estos valores y as� despu�s calcular el Saldo Inicial
'           Se cambia referencia de DAO 2.5-3.5 a DAO 3.6, con lo cual se pueden accedes las bases con Access 2000 y la funcionalidad que �ste ofrece. De todas maneras LPContab no cambia la base a Access 2000. Se puede seguir abriendo con 97
'13/09/2017 Se agranda campo para almacenar N�mero de Cartola, dado que se almacenaba en un byte y se aceptaban dos d�gitos a lo m�s. Ahora son 3 d�gitos y se almacena en un integer
'           En la exportaci�n a F-22 y F-29, se cambia la creaci�n de una base de datos por la copia de una base de datos vac�a que se instala en el actualizador, para evitar que la cree con Access 2000, producto del cambio a DAO 3.6
'           Se agrega reemplazo de CR y LF (cortes de l�nea) por nada en el ingreso de datos en los libros de compras, ventas y retenciones
'           Se agrega Saldo inicial a Libro de Caja Consolidado
'14/09/2017 Se elimina Configuraci�n de cuentas FUT y Exportaci�n a FUT, con lo cual se elimina ADO 2.8 (FireFox)
'15/09/2017 Se agrega al Administrador opci�n para exportar una empresa. �sta opci�n genera un ZIP con LPContab.mdb, el a�o seleccionado y el a�o anterior
'20/11/2017 Se agregan validaciones al documento asociado en el detale de un doc de compra o venta (notas de cr�dito o d�bito)
'           Se agrega Proporcionalidad de IVA a Reporte Supermercado
'           Se agrega mensaje de advertencia cuando un usuario modifica atributos de cuentas
'           Mejora en la validaci�n de IVA CF cuando el documento tiene m�s de dos meses
'           Se agrega nuevos r�gimenes tributarios: Renta Atribuida y Semi Integrado
'           Correcci�n de valores de NCC y NCF en el resumen de Supermercados
'           Validaci�n de fecha en el ingreso manual de registros en una cartola bancaria
'           No permitir ingresar datos en columna Otros Impuestos para el caso de VPE (tanto importaci�n como ingreso manual)
'           Se agrega la opci�n, en el Administrador, de crear dos usuarios Fiscalizadores que s�lo pueden ver informaci�n en el sistema. Tienen un per�odo en que est�n habilitados
'           Se modifica el importador de Entidades para que permita ingreso de RUT extranjeros
'           Importaci�n de entidades: ahora es obligatorio seleccionar el tipo de entidad: Cliente, Proveedor, socio, etc.
'           Se corrigen algunos nombres de comunas que cambiaron desde que se ingresaron la primara vez
'           Se agrega una opci�n de men� para modificar el comprobante de apertura Tributario, adem�s del Financiero
'           Se agrega opci�n para visualizar los planes de cuenta que provee el sistema, para revisarlos antes de decidir cu�l utilizar�.
'           Se modifica el proceso de centralizaci�n para que utilice como fecha del comprobante, el mes que se est� centralizando y no el �ltimo mes abierto.
'           Se ajustan los privilegios de los usuarios para las nuevas funcionalidades que se han agregado al sistema
'06/12/2017 Se agrega opci�n de Ajustes Extra Libro de Caja
'27/12/2017 Se agrega partida para una cuenta y la configuraci�n de partidas para las cuentas de los planes B�sico, Intermedio, Avanzado e IFRS
'28/12/2017 Se elimina recuadro 2 de From22 y se restingen los c�digos del Recuadro 3 de Fomr 22, a partir del a�o 2017
'04/01/2018 Se agrega opci�n de Base Imponible 14 TER
'10/01/2018 Se agrega importaci�n de Centro de Costo y �rea de Negocio a libros de Compras Ventas
'31/01/2018 Se agrega exportaci�n de declaraciones DJ 1923 Secci�n B y DJ 1924
'19/04/2018 Se agrega a la configuraci�n de una cuenta la opci�n de seleccionar una cuneta del Plan de Cuentas SII, para ser utilizada en la DJ 1847
'           Se agrega exportaci�n de declaraciones DJ 1847
'           Se ajusta el ingreso de fecha de emisi�n y RUT en los libros de compras y ventas
'           Se realiza un ajuste en Reporte de Activo Fijo Financiaro
'15/06/2018 Se agrega proceso de importaci�n de registro de compras SII
'           Se agrega configuraci�n de cuentas, �rea de negocio, centro de costo y proporcionalidad para cada proveedor, distinguiendo Afecto, Exento y Total
'04/07/2018 Se corrige error en Activo Fijo Financiero
'09/07/2018 Se agrega manual de M�dulo de Activo Fijo como opci�n de Men� Activo Fijo
'           Se genera versi�n 6.4.0
'25/07/2018 Se corrige error al seleccionar un libro en ventana de selecci�n de libros
'10/08/2018 Se corrige error en Ajustes Extra Libro de Caja en el c�lculo de
'              Ingresos devengados y no percibidos con  plazo mayor a  12 meses desde que  se emiti� dcto.
'              Ingresos devengados y no percibidos con plazo mayor a 12 meses desde que pago es  exigible
'14/08/2018 Se corrige error de duplicaci�n de documentos al importar en Libro de Caja (era porque se defini� el IdDoc como integere tabla LibroCaja)
'16/08/2018 Se corrige detalle en importaci�n desde Facturaci�n: faltaba marcar el campo EsTotalDoc cuando corresponde
'21/08/2018 Se cambia de Haber a Debe en la generaci�n de movimientos de Impuestos Adicionales: IVA retenido, en el caso de una NCC. Antes estaba s�lo para NCF
'           Se cambia link para Tutorial de LPContabilidad
'           Se agrega link para manuales de uso
'08/10/2018 Se corrige error al importar Registro Libro de Compras SII, identificadores de Area de Neg y Centro de Costo estaban al rev�s.
'10/10/2018 Se agrega ayuda de c�digos 14 TER a ventana de edici�n de cuenta
'12/10/2018 Se agrega opci�n de selecci�n m�ltiple de �rea de Negocio y Centro de Costo a balance clasificado y Estado de Resultado Clasificado
'16/10/2018 Se agrega impresi�n con papel foliado y firma en el pie de libros oficiales (nombre y rut contador) para Ajustes Extacontables y Base imponible
'           Se agrega opci�n de importar o no documentos y activos fijos pendientes del a�o anterior, en la opci�n "generar comprobante de Apertura"
'           Se agrega opci�n a info anal�tico para que incorpore o no los saldos de apertura
'17/10/2018 Se agrega mensaje para advertir a usuario cuando intenta modificar un documento que ya fue exportado al a�o siguiente
'18/10/2018 Se agrega listado de nuevos c�digos de Act. Econ�mica emitido por el SII. El sistema realiza la conversi�n autom�tica si existe
'29/10/2018 Se agregan nuevo reporte Balance Clasificado por �rea de Negocio y por Centro de Costo
'           Se agregan nuevo reporte Estado de Resultado por �rea de Negocio y por Centro de Costo
'           Se agrega manual a Importaci�n Registro de Compras SII
'31/10/2018 Se agrega nuevo reporte Libro Diario Esquem�tico
'19/12/2018 Se corrige error en FrmImportRemu cuando se pide detallado por centro de costo pero hay empreados sin centro de costo
'23/05/2019 Se agrega validaci�n a Config de Cuentas por Proveedor para Libro de Compras y Ventas, para que exija Centro de Costo y Area de Negocio si la cuenta as� lo indica en los atributos
'           Se agrega validaci�n a Importaci�n de Libro de Compras o Ventas (captura en el mismo libro) , para que exija Centro de Costo y Area de Negocio si la cuenta as� lo indica en los atributos
'           Se modifican algunos traspasos en 14 TER, desde Ajustes Extra Contable a Base Imponible
'27/05/2019 Se corrige validaci�n al capturar boletas de honorarios al libro de Retenciones
'28/05/2019 Se modifica texto de versi�n en ventana de inicio
'30/05/2019 Se corrige problema en resultado de IVA
'03/06/2019 Se modifica la ventana de IPC y factores por el tema del cambio de base del SII
'26/06/2019 Se corrige error al asignar las cuentas asociadas al proveedor en la importaci�n del Libro de Compras del SII
'03/07/2019 Se agrega a importaci�n de Configuraci�n de Cuentas por Proveedor del libro de Compras, que al crear la entidad si no exite, la marque de inmediato como proveedor
'           Se agrega a importaci�n de Configuraci�n de Cuentas por Cliente del libro de Ventas, que al crear la entidad si no exite, la marque de inmediato como cliente
'25/08/2019 Se modifica reporte de Activo Fijo de acuerdo a especificaciones de Victor Morales
'13/09/2019 Se modifican todos los enlaces con sistemas externos para que funciones en SLQ Server tambi�n
'           La importaci�n desde REmu puede ser con las 4 combinaciones: SQL Server - SQL Server, Access - Access y combinaciones de ambos
'16/09/2019 Se cambian t�tulos y textos de 14 TER a 14 TER A)
'           Se cambia help de importaci�n de comprobantes
'26/09/2019 Se cambia el esquema de traspaso a HR de la siguiente manera:
            'Formulario    Archivo generado     Carpeta                    Comentario
            '1. HR F29     F29_MMAA.mdb         HR\RUTS\NNNNNNNN\ImpConta
            '2. HR F22     F22_AA.mdb           HR\RUTS\NNNNNNNN\ImpConta
            '3. DJ 1879    DJ1879_MMAA.csv      HR\RUTS\NNNNNNNN\ImpConta  Si MM = 00 => todo el a�o
            '4. DJ 1923    DJ1923_AA.csv        HR\RUTS\NNNNNNNN\ImpConta
            '5. HRRAB      HRRAB_AA.csv         HR\RUTS\NNNNNNNN\ImpConta  Ac� no se invoca al wizard. Para invocarlo requiero c�digo
            '6. DJ 1924    DJ1924B_AA.csv       HR\RUTS\NNNNNNNN\ImpConta
            '              DJ1924C_AA.csv       HR\RUTS\NNNNNNNN\ImpConta
            '7. DJ 1947    DJ1947_AA.csv        HR\RUTS\NNNNNNNN\ImpConta
'26/09/2019 se soluciona error de generaci�n de comprobante de remuneraciones, que se presenta s�lo en SQL Server. Es porque no trunca glosa autom�ticamente.
'           Se usa SET ANSI_WARNINGS {ON|OFF} al inicio y fin de la funci�n
'26/09/2019 Se modifican t�tulos de listado de Ingresos y Egresos 14 TER
'           Se  agrega tipo de ajuste a c�lculo de razones financieras
'01/10/2019 Se corrige error al pagar un n�mero muy grande de documentos en un comprobante
'11/10/2019 Se agrega importaci�n de empresa desde HR tanto en Administrador como en Contabilidad
'14/10/2019 Se agrega importaci�n de empresas desde HR en el Administrador
'           Se agrega importaci�n de datos b�sicos de empresa en HR cuando se crea el priemr a�o de una empresa en Conta (si la empresa est� en HR)
'15/10/2019 Se usa SET ANSI_WARNINGS {ON|OFF} al inicio y fin de la funci�n de importaci�n de registro de compras SII
'08/11/2019 Se corrige error en c�lculo de Saldo Inicial en Libro de Caja Consolidado
'27/11/2019 Se corrige error en c�lculo Resumen de IVA
'10/12/2019 Se agrega Eliminar Empresa-A�o y Eliminar Empresa en Administrador, para el caso de SQL Server
'23/12/2019 Se agrega abrir un nuevo a�o en SQL Server desde una base de datos Access
'26/12/2019 Se agrega nuevo impuesto a las retenciones nacionales que corre desde enero 2020, con valor de 10,75%
'08/01/2020 Se corrige error que asignaba mal el a�o en el CorrigeBase a todas las tablas, cuando el corrige base se llamaba desde el a�o siguiente.
'15/01/2020 Se agrega Configuraci�n para Retiros y Dividendos y generaci�n de archivo RLI para HR-RAB
'27/01/2020 Se agrega c�lculo de Saldos a otros documentos, basado en el Total de documento, como saldo inicial y los movimientos en comprobantes a los cuales el documento est� asociado
'           Se agrega opci�n para traer "Otros Documentos" desde a�o anterior, manteniendo el Saldo
'07/02/2020 Se corrige error en nueva entidad cunado no es Rut V�lido
'14/02/2020 Se agrega JoinEmpAno en RecalSaldos, GenResOImp y GenResOImpRecup
'03/03/2020 Se agrega funcionalidad de RLI HR RAB: Configuraci�n de Cuentas y generaci�n de archivo a trav�s de la ventana de Ajustes Extra Contables RLI - RAB
'09/03/2020 Se agrega al saldo inicial de Ajustes Extra Contables 14 TER el saldo inicial de otros ingresos
'           Se garegan mensajes de la ley 21210 para indicar que estamos trabajando en el tema
'13/03/2020 Se corrige error al tomar los datos de una empresa desde HR (tipo inv�lido)
'17/03/2020 Se corrige la funci�n para exportar e importar un plan de cuentas para que generen y lean respectivamente el mismo formato de archivo
'31/03/2020 Se corrige error de tipo de dato al obtener los datos de una empresa nueva desde HR
'01/04/2020 Se agrega funcionalidad de obtener factores de correcci�n monetaria (actualizaci�n) desde el SII pata todo el a�o (matriz completa)
'           Se usan estos factores en el Reporte de Activo Fijo Tributario, para obtener el factor de actualizaci�n de cada activo fijo, comprado el a�o actual, dependiendo de la fecha de compra y de la fecha del informe
'           Se modifica funci�n que genera DJ1924, para que ponga cero en los valores que no se generan en la funci�n, de acuerdo a lo solicitado por Hugo Lillo
'06/04/2020 Se modifica ventana de configuraci�n de impuestos para que acepte decimales en Cr�dito Art. 33 Bis desde 2015 en adelante
'13/04/2020 Se agrega a la exportaci�n/importaci�n del plan de cuentas, �rea Cuentas B�sicas, el campo ParamEmpresa.Codigo para usarlo en el caso de la config. de cuentas de remuneraciones.
'           Se agrega ORDER BY a exportaci�n de Plan de Cuentas, �rea Cuentas B�sicas para facilitar la lectura del archivo de exportaci�n
'16/04/2020 Se modifica Edici�n de Monedas para que no se modifiquen las monedas fijas, predefinidas por el sistema
'23/04/2020 Se agrega Capital Propio a Exportaci�n a F22
'03/07/2020 Se agrega nueva depreciaci�n Ley 21.210 Activos Fijos
'07/08/2020 Se agrega importaci�n Registro de Ventas y Configuraci�n detallada de cuentas por cliente para libro de Ventas
'27/08/2020 Se agregan registros adicionales con Otros Impuestos codificados al Registro de Ventas
'01/09/2020 Se incorporan mejoras a importaci�n Registro Ventas SII, incluyendo documento de referencia
'09/09/2020 Se corrige modifica �ndice de ImpAdic en Access para que ioncluya empresa-a�o
'           Se agrega copia de Cuentas Ajuste Extra Contables y Cuentas Ajustes Extra Contables RLI de un a�o a otros en Access y SQL Server, incluso cuando se genera nuevo a�o en SQL Server desde Access
'01/10/2020 Se agrega Capital Propio Simplificado al Sistema con todas las ventanas adjuntas
'           Se agrega validaci�n de tabla Adm_Region_Contrib de HR para que no de error cuando no est� a�n en la base de datos de HR
'02/10/2020 Se corrige error en Resumen de IVA. No se considerana m�s de un IVA Ret Parcial o Total
'14/10/2020 Se agrega doc de Remuneraciones (LIB_REMU) a Anal�tico por Cuenta y por Entidad
'           Se agrega LIB_REMU a otras funcionalidades como generar docs pendientes a�oa anterior, que s�lo consideraba LIB_OTROS
'16/10/2020 Se modifica importaci�n de compras y ventas del SII para que tome la configuraci�n de impuestos adicionales del usuario. Estaba tomamdo la por omisi�n
'           Se agrega doc asociado a importaci�n de comporbantes, para que se pueda asociar un documento a un registro de un comprobante. El documento debe ya existir en el sistema
'19/10/2020 Se restringe la validaci�n del DL 825 en cuanto a plazo de IVA en libro de compras y ventas, s�lo a las compras
'21/10/2020 Se corrige error en importaci�n de registro de compras SII, faltaba limpiar variables de doc asociado al cambiar de documento
'02/11/2020 Se agrega Capital Propio Simplificado por Variaci�n Anual y se agregarn algunos elementos al CPS general
'12/11/2020 Se modifica fecha t�rmino Depreciaci�n Ley 21.210
'           Se agrega Depreciaci�n Ley 21.256 Art 3
'24/11/2020 Se corrige error en nombre de campo de tabla EmpresasAno al traer Activos Fijos del a�o anterior
'30/11/2020 Se agregan cuentas a Plan de Cuentas SII que son v�lidas desde el 2020 (por eso se agrega campo A�oDesde a tabla PlanCuentasSII)
'14/12/2020 Se agrega nueva depreciaci�n Ley 21256
'27/12/2020 Se hacen cambios en pantallas por cambios en DDJJ 1847
'04/01/2020 Se agrega un tipo de cuenta m�s a la configuraci�n de remuneraciones: Diferencias por Cobrar
'           Se agregan/eliminan elementos a los Ajustes Extra Contables, que corresponden a mofificaciones tributarias a�o 2020
'05/01/2021 Se agrega a la importaci�n de Comprobantes, la opci�n de asociar a un movimiento de un comprobante, un documento de tipo OTROS DOCS.
'15/01/2021 Se agregan elementos para 14D a ajustes extracontables
'21/01/2021 Se agrega Base Imponible 14D
'           Se agrega restricci�n a entidad relacionada, s�lo si es del tipo Art. 14 A Regimen Semi Integrado
'03/03/2021 Se agrega traspaso de franquicias nuevas como 14 A cuando se crea nuevo a�o
'           Se agregan nuevos c�digos de partidas para r�gimen 14 A
'10/03/2021 Se agrega exportaci�n de Capital Propio Simplificado - CPS a HR RAD
'12/03/2021 Se modifica Capital Propio Simplificado de acuerdo a nuevas instrucciones SII
'30/03/2021 Se agrega importaci�n de Otros Documentos
'           Se agrega opci�n que permnite revisar que el detalle de los documentos est� cuadrado (esto deber�a estar siempre cuadrado pero si la base est� da�ada pueden eliminarse registros sin que el cliente se de cuenta y generar errores inesperados).
'02/05/2021 Se cambia c�digo de partida en DJ1923, DJ1847 y RAB para que del 2019 hacia atr�s use el �ndice de la lista en vez del c�digo del SII, como se hace desde el 2020 (inconsistencia)
'           Cambios en columnas de libro de caja, ingresos y egresos por modificaciones SII
'07/05/2021 Se agrega bot�n de impresi�n resumida en Libro de Caja
'25/07/2021 Se agregan modificaciones a libro de caja ingresos y egresos y nueva columna Monto que Afecta a Base Imponible

