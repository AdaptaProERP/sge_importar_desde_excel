// Programa   : CSLIBROSCOM30 (CSLIBROS 3.0 Libro de Compras)
// Fecha/Hora : 30/12/2007 16:35:18
// PropÃ³sito : Formulario de Emisión de Libro de Compras con salida a Excel
// Aplicación : 
// Tipo       : STD00000


#include "DPXBASE.CH"

PROCE MAIN(dDesde,dHasta,lGet)
  LOCAL oBtn,oFont,oData,nClrText:=CLR_BLUE,oTable,nAt,aForms
  LOCAL cMemo:=""

  DEFAULT dDesde:=FCHINIMES(oDp:dFecha),;
          dHasta:=FCHFINMES(oDp:dFecha)


  // Contribuyente Especial
  IF lGet=NIL .AND. oDp:lConEsp
     lGet:=.F.
  ENDIF

  DEFAULT lGet:=.T.


//  DEFAULT lConEsp:=.F., lPlanilla:=.F.

  aForms :={}
  AADD(aForms,"Formato Separado de Documentos")
  AADD(aForms,"Formato Columna Unica")
  AADD(aForms,"Providencia Administrativa SNAT/2002/1.445")
  AADD(aForms,"Formato de Columnas Separadas de IVA")
  AADD(aForms,"Formato Col. Sep. de IVA Decreto 3085")

  oData:=DATASET("CSLIBROS","ALL")

  oCom:=DPEDIT():New("CSLIBROS 3.0 para AdaptaPro. Emisión Libro de Compras","CSLIBROSCOM.edt","oCom",.T.)
  oCom:cFileChm     :=""
  oCom:cTopic       :=""

  oCom:dDesde       :=dDesde // FCHINIMES(oDp:dFecha)
  oCom:dHasta       :=dHasta // FCHFINMES(oDp:dFecha)

  oCom:lCSombCab :=oData:Get("lCSombCab" ,.T.)
  oCom:lCSombTot :=oData:Get("lCSombTot" ,.T.)
  oCom:nCLenRs   :=oData:Get("nCLenRs"   ,40)
  oCom:nCPorAju  :=oData:Get("nCPorAju"  ,100)
  oCom:lCSepAju  :=oData:Get("lCSepAju"  ,.T.)
  oCom:lCSepMil  :=oData:Get("lCSepMil"  ,.T.)
  oCom:lCRetAju  :=oData:Get("lCRetAju"  ,.T.)
  oCom:lCDocAju  :=oData:Get("lCDocAju"  ,.T.)
  oCom:lCHojAju  :=oData:Get("lCHojAju"  ,.T.)
  oCom:lCSepIva  :=oData:Get("lCSepIva"  ,.T.)
  oCom:lCRifCab  :=oData:Get("lCRifCab"  ,.T.)
  oCom:lCNumPag  :=oData:Get("lCNumPag"  ,.T.)
  oCom:lCConDin  :=oData:Get("lCConDin"  ,.T.)
  oCom:nCForm    :=oData:Get("nCForm"    ,1)

****AG20080304
  oCom:lCLibRti  :=oData:Get("lCLibRti"  ,.F.)
****AG20080304

  oCom:lEDetIdio  :=oData:Get("lEDetIdio"  ,.T.)
  oCom:lEDetSep   :=oData:Get("lEDetSep"   ,.T.)
  oCom:cEFunCel   :=oData:Get("cEFunCel"   ,SPACE(10))
  oCom:cEFunSum   :=oData:Get("cEFunSum"   ,SPACE(10))
  oCom:cEFunSi    :=oData:Get("cEFunSi"    ,SPACE(10))
  oCom:cESepMil   :=oData:Get("cESepMil"   ," ")
  oCom:cESepDec   :=oData:Get("cESepDec"   ," ")
  oCom:cESepLis   :=oData:Get("cESepLis"   ," ")

  oData:End(.F.)

  oCom:aMeses :={"Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"} 
  oCom:nMes   :=MONTH(dDesde)
  oCom:nAno   :=YEAR(dDesde)
  oCom:nRecord:=0
  oCom:cCodSuc:=oDp:cSucursal
*  oCom:lConEsp:=lConEsp
  oCom:dFecha :=oDp:dFecha
  oCom:aForms :=aForms
  oCom:lGet   :=lGet
  oCom:cMemo  :="Esta funcionalidad esta en fase revisión y Actualización,"+CRLF+;
                "aún no esta en producción, agradecemos su validación."

  @ 6.8, 1.0 FOLDER oCom:oFolder ITEMS "Libro","Valores","Excel"

  SETFOLDER( 1)

  @ 3,2 SAY "Año:" RIGHT
  @ 3,2 SAY "Mes:" RIGHT
  @ 3,1 SAY GetFromVar("{oDp:xDPSUCURSAL}")+":"
  @ 4,1 SAY "Forma Excel"

  @ 0.5,3 GET oCom:nAno PICTURE "9999" SPINNER WHEN oCom:lGet;
          VALID oCom:HACERFECHA()

  @ 2.0,3 COMBOBOX oCom:oMes VAR oCom:nMes ITEMS oCom:aMeses WHEN oCom:lGet;
          VALID oCom:HACERFECHA()

  @ 3,2 SAY oCom:oSayRecord PROMPT "Registros:"

  // Forma Excel
  @ 0,1 COMBOBOX oCom:oForm VAR oCom:nCForm ITEMS oCom:aForms ON CHANGE oCom:ChangeControls()

  ComboIni(oCom:oForm)

  oCom:oForm:cMsg    :="Forma Excel"
  oCom:oForm:cTooltip:="Forma Excel"

  // SUCURSAL
  @ .1,06 BMPGET oCom:oCodSuc VAR oCom:cCodSuc;
                 VALID CERO(oCom:cCodSuc,NIL,.T.) .AND.;
                            oCom:FindCodSuc();
                 NAME "BITMAPS\FIND.BMP"; 
                 ACTION (oDpLbx:=DpLbx("DPSUCURSAL",NIL,NIL),;
                         oDpLbx:GetValue("SUC_CODIGO",oCom:oCodSuc)); 
                 SIZE 48,10

  @ 3,2 SAY oCom:oSucNombre PROMPT SQLGET("DPSUCURSAL","SUC_DESCRI","SUC_CODIGO"+GetWhere("=",oCom:cCodSuc));
            UPDATE

  @ 3,5 GET oCom:cMemo MULTILINE READONLY

  @ 02,01 METER oCom:oMeter VAR oCom:nRecord


  @ 05,20 BMPGET oCom:oDesde  VAR oCom:dDesde;
                  PICTURE "99/99/9999";
                  NAME "BITMAPS\Calendar.bmp";
                  ACTION LbxDate(oCom:oDesde ,oCom:dDesde);
                  SIZE 76,20;
                  WHEN .t.;
                  FONT oFont

   oCom:oDesde:cToolTip:="F6: Calendario"

   @ 05,60     BMPGET oCom:oHasta  VAR oCom:dHasta;
               PICTURE "99/99/9999";
               NAME "BITMAPS\Calendar.bmp";
               ACTION LbxDate(oCom:oHasta,oCom:dHasta);
               SIZE 76,20;
               WHEN .t.;
               FONT oFont

  SETFOLDER( 2)

  @ 1,1 CHECKBOX oCom:lCSombCab PROMPT ANSITOOEM("Sombrear Cabecera")
  @ 2,1 CHECKBOX oCom:lCSombTot PROMPT ANSITOOEM("Sombrear Totales")
  @ 3,1 GET oCom:oNCLenRs  VAR oCom:nCLenRs
  oCom:oNCLENRS:cMsg    :="Indique Longitud"
  oCom:oNCLENRS:cToolTip:="Indique Longitud"
  @ oCom:oNCLENRS:nTop-08,oCom:oNCLENRS:nLeft SAY "Longitud de Razón Social" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 4,1 GET oCom:oNCPorAju VAR oCom:nCPorAju
  oCom:oNCPORAJU:cMsg    :="Indique Porcentaje"
  oCom:oNCPORAJU:cToolTip:="Indique Porcentaje"
  @ oCom:oNCPORAJU:nTop-08,oCom:oNCPORAJU:nLeft SAY "Porcentaje Ajuste Tamaño" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 


  @ 5,1 CHECKBOX oCom:oCSepAju VAR oCom:lCSepAju PROMPT ANSITOOEM("Separar Ajustes en un Cuadro Distinto") ON CHANGE oCom:ChangeControls() WHEN oCom:nCForm<>2
  @ 6,1 CHECKBOX oCom:lCSepMil PROMPT ANSITOOEM("Incluir Separadores de Mil")
  @ 7,1 CHECKBOX oCom:oCRetAju VAR oCom:lCRetAju PROMPT ANSITOOEM("Colocar Retenciones Fuera del Perí­odo en Ajustes") WHEN oCom:lCSepAju .AND. oCom:nCForm<>2
  @ 8,1 CHECKBOX oCom:oCDocAju VAR oCom:lCDocAju PROMPT ANSITOOEM("Colocar Documentos Fuera del Perí­odo en Ajustes") WHEN oCom:lCSepAju .AND. oCom:nCForm<>2
  @ 9,1 CHECKBOX oCom:oCHojAju VAR oCom:lCHojAju PROMPT ANSITOOEM("Colocar Ajustes en la misma Hoja de Cálculo") WHEN oCom:lCSepAju .AND. oCom:nCForm<>2
  @ 10,1 CHECKBOX oCom:oCSepIva VAR oCom:lCSepIva PROMPT ANSITOOEM("Separar en Filas distintas la alícuota general y la reducida")
  @ 11,1 CHECKBOX oCom:lCRifCab PROMPT ANSITOOEM("Agregar Rif en Cabecera")
  @ 12,1 CHECKBOX oCom:lCNumPag PROMPT ANSITOOEM("Numerar Páginas")
  @ 13,1 CHECKBOX oCom:lCConDin PROMPT ANSITOOEM("Consecutivo Dinámico")
     
  SETFOLDER( 3)

  @ 1,1 CHECKBOX oCom:lEDetIdio PROMPT ANSITOOEM("Detección Automática de Idioma") ON CHANGE oCom:ChangeControls()
  @ 2,1 CHECKBOX oCom:lEDetSep  PROMPT ANSITOOEM("Detección Automática de Separadores") ON CHANGE oCom:ChangeControls()
  @ 4,1 GET oCom:oCEFUNCEL  VAR oCom:cEFunCel WHEN !oCom:lEDetIdio
  oCom:oCEFUNCEL:cMsg     :="Indique la como se describe literalmente la función"
  oCom:oCEFUNCEL:cToolTip :="Indique la como se describe literalmente la función"
  @ oCom:oCEFUNCEL:nTop-08,oCom:oCEFUNCEL:nLeft SAY "Función Celda" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 5,1 GET oCom:oCEFUNSUM  VAR oCom:cEFunSum WHEN !oCom:lEDetIdio
  oCom:oCEFUNSUM:cMsg     :="Indique la como se describe literalmente la función"
  oCom:oCEFUNSUM:cToolTip :="Indique la como se describe literalmente la función"
  @ oCom:oCEFUNSUM:nTop-08,oCom:oCEFUNSUM:nLeft SAY "Función Suma" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 6,1 GET oCom:oCEFUNSI  VAR oCom:cEFunSi WHEN !oCom:lEDetIdio
  oCom:oCEFUNSI:cMsg     :="Indique la como se describe literalmente la función"
  oCom:oCEFUNSI:cToolTip :="Indique la como se describe literalmente la función"
  @ oCom:oCEFUNSI:nTop-08,oCom:oCEFUNSI:nLeft SAY "Función Si" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 8,1 GET oCom:oCESEPMIL VAR oCom:cESepMil WHEN !oCom:lEDetSep
  oCom:oCESEPMIL:cMsg     :="Indique el caracter separador"
  oCom:oCESEPMIL:cToolTip :="Indique el caracter separador"
  @ oCom:oCESEPMIL:nTop-08,oCom:oCESEPMIL:nLeft SAY "Separador Miles" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 9,1 GET oCom:oCESEPDEC VAR oCom:cESepDec WHEN !oCom:lEDetSep
  oCom:oCESEPDEC:cMsg     :="Indique el caracter separador"
  oCom:oCESEPDEC:cToolTip :="Indique el caracter separador"
  @ oCom:oCESEPDEC:nTop-08,oCom:oCESEPDEC:nLeft SAY "Separador Decimales" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 10,1 GET oCom:oCESEPLIS VAR oCom:cESepLis WHEN !oCom:lEDetSep
  oCom:oCESEPLIS:cMsg     :="Indique el caracter separador"
  oCom:oCESEPLIS:cToolTip :="Indique el caracter separador"
  @ oCom:oCESEPLIS:nTop-08,oCom:oCESEPLIS:nLeft SAY "Separador Lista" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  SETFOLDER(0)

  oCom:Activate({||oCom:INICIO()})

RETURN oCom

/*
// Grabar Empresa
*/
FUNCTION LibCompra()
  LOCAL oData,nAt

*  CursorWait()
 
  oData:=DATASET("CSLIBROS","ALL")

  oData:Set("lCSombCab" ,oCom:lCSombCab)
  oData:Set("lCSombTot" ,oCom:lCSombTot)
  oData:Set("nCLenRs"   ,oCom:nCLenRs)
  oData:Set("nCPorAju"  ,oCom:nCPorAju)
  oData:Set("lCSepAju"  ,oCom:lCSepAju)
  oData:Set("lCSepMil"  ,oCom:lCSepMil)
  oData:Set("lCRetAju"  ,oCom:lCRetAju)
  oData:Set("lCDocAju"  ,oCom:lCDocAju)
  oData:Set("lCHojAju"  ,oCom:lCHojAju)
  oData:Set("lCSepIva"  ,oCom:lCSepIva)
  oData:Set("lCRifCab"  ,oCom:lCRifCab)
  oData:Set("lCNumPag"  ,oCom:lCNumPag)
  oData:Set("lCConDin"  ,oCom:lCConDin)
  oData:Set("nCForm"    ,oCom:nCForm)

  oData:Set("lEDetIdio"  ,oCom:lEDetIdio)
  oData:Set("lEDetSep"   ,oCom:lEDetSep)
  oData:Set("cEFunCel"   ,oCom:cEFunCel)
  oData:Set("cEFunSum"   ,oCom:cEFunSum)
  oData:Set("cEFunSi"    ,oCom:cEFunSi)
  oData:Set("cESepMil"   ,oCom:cESepMil)
  oData:Set("cESepDec"   ,oCom:cESepDec)
  oData:Set("cESepLis"   ,oCom:cESepLis)

  oData:Save()
  oData:End(.t.)

  EJECUTAR("CSLIBROSCOM3001",oCom) // Todo se unifica en un solo proceso
//  MsgInfo("El programa CSLIBROSCOM3001 está en desarrollo","Información")

RETURN .T.

FUNCTION FINDCODSUC()
  oCom:oSucNombre:Refresh(.T.) 

  IF !oCom:cCodSuc==SQLGET("DPSUCURSAL","SUC_CODIGO","SUC_CODIGO"+GetWhere("=",oCom:cCodSuc))
    EVAL(oCom:oCodSuc:bAction)
    RETURN .F.
  ENDIF
RETURN .T.

FUNCTION ChangeControls()
IF oCom:nCForm=2
   oCom:lCSepAju:=.F.
   oCom:oCSepAju:Refresh()
ENDIF
IF !oCom:lCSepAju
   oCom:lCRetAju:=.F.
   oCom:oCRetAju:Refresh()
   oCom:lCHojAju:=.F.
   oCom:oCHojAju:Refresh()
   oCom:lCDocAju:=.F.
   oCom:oCDocAju:Refresh()
ENDIF
IF oCom:lEDetIdio
   oCom:cEFunCel:=SPACE(10)
   oCom:oCEFUNCEL:Refresh()
   oCom:cEFunSum:=SPACE(10)
   oCom:oCEFUNSUM:Refresh()
   oCom:cEFunSi:=SPACE(10)
   oCom:oCEFUNSi:Refresh()
ENDIF
IF oCom:lEDetSep
   oCom:cESepMil:=" "
   oCom:oCESEPMIL:Refresh()
   oCom:cESepDec:=" "
   oCom:oCESEPDEC:Refresh()
   oCom:cESepLis:=" "
   oCom:oCESEPLIS:Refresh()
ENDIF
RETURN .T.


FUNCTION INICIO()
   LOCAL oCursor,oBar,oBtn,oFont,oCol
   LOCAL oDlg:=oCom:oDlg
   LOCAL nLin:=0

   DEFINE CURSOR oCursor HAND
   DEFINE BUTTONBAR oBar SIZE 52+15,60-15+15 OF oDlg 3D CURSOR oCursor
   DEFINE FONT oFont  NAME "Tahoma"   SIZE 0, -14 BOLD


   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\RUN.BMP",NIL,"BITMAPS\RUNG.BMP";
          TOP PROMPT "Ejecutar"; 
          ACTION (CursorWait(),;
                  oCom:LIBCOMPRA(),;
                  EJECUTAR("IVALOAD",oCom:dFecha))

   oBtn:cToolTip:="Guardar"

   oCom:oBtnRun:=oBtn


   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\XSALIR.BMP";
          TOP PROMPT "Cerrar"; 
          ACTION (oCom:Cancel()) CANCEL

   oBar:SetColor(CLR_BLACK,oDp:nGris)

   AEVAL(oBar:aControls,{|o,n| o:SetColor(CLR_BLACK,oDp:nGris) })

RETURN .T.

FUNCTION HACERFECHA()

  oCom:dDesde:=CTOD("01/"+STRZERO(oCom:oMes:nAt,2)+"/"+STRZERO(oCom:nAno))
  oCom:dHasta:=FCHFINMES(oCom:dDesde)

  oCom:oDesde:VarPut(oCom:dDesde,.T.)
  oCom:oHasta:VarPut(oCom:dHasta,.T.)

RETURN .T.





