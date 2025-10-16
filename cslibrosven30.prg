// Programa   : CSLIBROSVEN30 (CSLIBROS 3.0 Libro de Ventas)
// Fecha/Hora : 29/11/2007 17:36:43
// Prop�sito  : Formulario de Emisi�n de Libro de Ventas con salida a Excel
// Aplicaci�n : 
// Tipo       : STD00000
// Creado Por : Daniel Ram�rez
// Observaci�n: 
// Modificaciones: (DR20110915a) Se agrega par�metro para incluir Retenciones con fecha distintas a la de la factura
//                               dentro del grupo "Fuera del Per�odo"

#include "DPXBASE.CH"

PROCE MAIN(dDesde,dHasta)
  LOCAL oBtn,oFont,oData,nClrText:=CLR_BLUE,oTable,nAt,cModelo:="",aSeries:={},aForms

  DEFAULT dDesde:=FCHINIMES(oDp:dFecha),;
          dHasta:=FCHFINMES(oDp:dFecha)

//  DEFAULT lConEsp:=.F., lPlanilla:=.F.

  aSeries:=ASQL("SELECT SFI_MODELO FROM DPSERIEFISCAL")
  cModelo:=SQLGET("DPTIPDOCCLI","TDC_SERIEF","TDC_TIPO='FAV'")
  aForms :={}
  AADD(aForms,"Formato Separado de Documentos")
  AADD(aForms,"Formato Columna Unica")
  AADD(aForms,"Providencia Administrativa SNAT/2002/1.445")
  AADD(aForms,"Formato de Columnas Separadas de IVA")

  AEVAL(aSeries,{|a,n|aSeries[n]:=a[1]})
  AADD(aSeries,"Todos")

  oData:=DATASET("CSLIBROS","ALL")

  oVen:=DPEDIT():New("CSLIBROS 3.0 para AdaptaPro. Emisi�n Libro de Ventas","CSLIBROSVEN.edt","oVen",.T.)
  oVen:cFileChm     :=""
  oVen:cTopic       :=""

  oVen:dDesde       :=dDesde // FCHINIMES(oDp:dFecha)
  oVen:dHasta       :=dHasta // FCHFINMES(oDp:dFecha)

  oVen:lVControl :=oData:Get("lVControl" ,.T.)
  oVen:lVSombCab :=oData:Get("lVSombCab" ,.T.)
  oVen:lVSombTot :=oData:Get("lVSombTot" ,.T.)
  oVen:nVLenRs   :=oData:Get("nVLenRs"   ,40)
  oVen:nVPorAju  :=oData:Get("nVPorAju"  ,100)
  oVen:lVSepAju  :=oData:Get("lVSepAju"  ,.T.)
  oVen:lVRetFec  :=oData:Get("lVRetFec"  ,.F.) // DR20110915a. Incluye Retenciones Fuera de Fecha como "Fuera del Per�odo"
  oVen:lVSepMil  :=oData:Get("lVSepMil"  ,.T.)
  oVen:lVRetAju  :=oData:Get("lVRetAju"  ,.T.)
  oVen:lVRifCab  :=oData:Get("lVRifCab"  ,.T.)
  oVen:lVNumPag  :=oData:Get("lVNumPag"  ,.T.)
  oVen:lVConDin  :=oData:Get("lVConDin"  ,.T.)
  oVen:nVForm    :=oData:Get("nVForm"    ,1)
  oVen:lVColRti  :=oData:Get("lVColRti"  ,.F.)
  oVen:lVLibRti  :=oData:Get("lVLibRti"  ,.F.)

  oVen:lEDetIdio  :=oData:Get("lEDetIdio"  ,.T.)
  oVen:lEDetSep   :=oData:Get("lEDetSep"   ,.T.)
  oVen:cEFunCel   :=oData:Get("cEFunCel"   ,SPACE(10))
  oVen:cEFunSum   :=oData:Get("cEFunSum"   ,SPACE(10))
  oVen:cEFunSi    :=oData:Get("cEFunSi"    ,SPACE(10))
  oVen:cESepMil   :=oData:Get("cESepMil"   ," ")
  oVen:cESepDec   :=oData:Get("cESepDec"   ," ")
  oVen:cESepLis   :=oData:Get("cESepLis"   ," ")

  oData:End(.F.)

  oVen:aMeses :={"Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"} 
  oVen:nMes   :=MONTH(oDp:dFecha)
  oVen:nAno   :=YEAR(oDp:dFecha)
  oVen:nRecord:=0
  oVen:cCodSuc:=oDp:cSucursal
  oVen:cModelo:=cModelo
*  oVen:lConEsp:=lConEsp
  oVen:dFecha :=oDp:dFecha
  oVen:aForms :=aForms

  @ 6.8, 1.0 FOLDER oVen:oFolder ITEMS "Libro","Valores","Excel"

  SETFOLDER( 1)

  @ 3,2 SAY "A�o:" RIGHT
  @ 3,2 SAY "Mes:" RIGHT
  @ 3,1 SAY "Forma Fiscal:" 
  @ 4,1 SAY GetFromVar("{oDp:xDPSUCURSAL}")+":"
  @ 5,1 SAY "Forma Excel"

  @ 0.5,3 GET oVen:nAno PICTURE "9999" SPINNER

  @ 2.0,3 COMBOBOX oVen:oMes VAR oVen:nMes ITEMS oVen:aMeses

  @ 3,2 SAY oVen:oSayRecord PROMPT "Registros:"

  // Modelos
  @ 0,1 COMBOBOX oVen:oModelo VAR oVen:cModelo ITEMS aSeries

  ComboIni(oVen:oModelo)

  oVen:oModelo:cMsg    :="Serie Fiscal"
  oVen:oModelo:cTooltip:="Serie Fiscal"

  // Forma Excel
  @ 0,1 COMBOBOX oVen:oForm VAR oVen:nVForm ITEMS oVen:aForms ON CHANGE oVen:ChangeControls()

  ComboIni(oVen:oForm)

  oVen:oForm:cMsg    :="Forma Excel"
  oVen:oForm:cTooltip:="Forma Excel"

  // SUCURSAL
  @ .1,06 BMPGET oVen:oCodSuc VAR oVen:cCodSuc;
                 VALID CERO(oVen:cCodSuc,NIL,.T.) .AND.;
                            oVen:FindCodSuc();
                 NAME "BITMAPS\FIND.BMP"; 
                 ACTION (oDpLbx:=DpLbx("DPSUCURSAL",NIL,NIL),;
                         oDpLbx:GetValue("SUC_CODIGO",oVen:oCodSuc)); 
                 SIZE 48,10

  @ 3,2 SAY oVen:oSucNombre PROMPT SQLGET("DPSUCURSAL","SUC_DESCRI","SUC_CODIGO"+GetWhere("=",oVen:cCodSuc));
            UPDATE

  @ 02,01 METER oVen:oMeter VAR oVen:nRecord

  SETFOLDER( 2)

  @ 1,1 CHECKBOX oVen:oVControl VAR oVen:lVControl PROMPT ANSITOOEM("Tomar N�mero de Factura cuando este vac�o el n�mero fiscal") FONT oFont COLOR nClrText,NIL WHEN oVen:nVForm<>2
  @ 2,1 CHECKBOX oVen:lVSombCab PROMPT ANSITOOEM("Sombrear Cabecera")
  @ 3,1 CHECKBOX oVen:lVSombTot PROMPT ANSITOOEM("Sombrear Totales")
  @ 4,1 GET oVen:oNVLenRs  VAR oVen:nVLenRs
  oVen:oNVLENRS:cMsg    :="Indique Longitud"
  oVen:oNVLENRS:cToolTip:="Indique Longitud"
  @ oVen:oNVLENRS:nTop-08,oVen:oNVLENRS:nLeft SAY "Longitud de Raz�n Social" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 5,1 GET oVen:oNVPorAju VAR oVen:nVPorAju
  oVen:oNVPORAJU:cMsg    :="Indique Porcentaje"
  oVen:oNVPORAJU:cToolTip:="Indique Porcentaje"
  @ oVen:oNVPORAJU:nTop-08,oVen:oNVPORAJU:nLeft SAY "Porcentaje Ajuste Tama�o" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 


  @ 6,1 CHECKBOX oVen:oVSepAju VAR oVen:lVSepAju PROMPT ANSITOOEM("Separar Ajustes en un Cuadro Distinto") ON CHANGE oVen:ChangeControls() WHEN oVen:nVForm<>2
  @ 7,1 CHECKBOX oVen:lVSepMil PROMPT ANSITOOEM("Incluir Separadores de Mil")
  @ 8,1 CHECKBOX oVen:oVRetAju VAR oVen:lVRetAju PROMPT ANSITOOEM("Colocar Retenciones Fuera del Per�odo en Ajustes") WHEN oVen:lVSepAju .AND. oVen:nVForm<>2 .AND. !oVen:lVLibRti
  @ 9,1 CHECKBOX oVen:lVRifCab PROMPT ANSITOOEM("Agregar Rif en Cabecera")
  @ 10,1 CHECKBOX oVen:lVNumPag PROMPT ANSITOOEM("Numerar P�ginas")
  @ 11,1 CHECKBOX oVen:lVConDin PROMPT ANSITOOEM("Consecutivo Din�mico")
  @ 12,1 CHECKBOX oVen:lVColRti PROMPT ANSITOOEM("Agregar Columna con N�mero de Comprobante Retenci�n")
  @ 13,1 CHECKBOX oVen:oVLibRti VAR oVen:lVLibRti PROMPT ANSITOOEM("Colocar Retenciones fuera del Per�odo en el Libro de Ventas") ON CHANGE oVen:ChangeControls() WHEN oVen:nVForm<>2 .AND. !oVen:lVRetAju
  // DR20110915a. Incluir Retenciones fuera de Fecha como "Fuera del Per�odo"
  @ 14,1 CHECKBOX oVen:oVRetFec VAR oVen:lVRetFec PROMPT ANSITOOEM("Colocar Retenciones con fecha distinta a Factura como 'Fuera del Per�odo'") WHEN oVen:lVSepAju .AND. oVen:nVForm<>2
     
  SETFOLDER( 3)

  @ 1,1 CHECKBOX oVen:lEDetIdio PROMPT ANSITOOEM("Detecci�n Autom�tica de Idioma") ON CHANGE oVen:ChangeControls()
  @ 2,1 CHECKBOX oVen:lEDetSep  PROMPT ANSITOOEM("Detecci�n Autom�tica de Separadores") ON CHANGE oVen:ChangeControls()
  @ 4,1 GET oVen:oCEFUNCEL  VAR oVen:cEFunCel WHEN !oVen:lEDetIdio
  oVen:oCEFUNCEL:cMsg     :="Indique la como se describe literalmente la funci�n"
  oVen:oCEFUNCEL:cToolTip :="Indique la como se describe literalmente la funci�n"
  @ oVen:oCEFUNCEL:nTop-08,oVen:oCEFUNCEL:nLeft SAY "Funci�n Celda" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 5,1 GET oVen:oCEFUNSUM  VAR oVen:cEFunSum WHEN !oVen:lEDetIdio
  oVen:oCEFUNSUM:cMsg     :="Indique la como se describe literalmente la funci�n"
  oVen:oCEFUNSUM:cToolTip :="Indique la como se describe literalmente la funci�n"
  @ oVen:oCEFUNSUM:nTop-08,oVen:oCEFUNSUM:nLeft SAY "Funci�n Suma" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 6,1 GET oVen:oCEFUNSI  VAR oVen:cEFunSi WHEN !oVen:lEDetIdio
  oVen:oCEFUNSI:cMsg     :="Indique la como se describe literalmente la funci�n"
  oVen:oCEFUNSI:cToolTip :="Indique la como se describe literalmente la funci�n"
  @ oVen:oCEFUNSI:nTop-08,oVen:oCEFUNSI:nLeft SAY "Funci�n Si" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 8,1 GET oVen:oCESEPMIL VAR oVen:cESepMil WHEN !oVen:lEDetSep
  oVen:oCESEPMIL:cMsg     :="Indique el caracter separador"
  oVen:oCESEPMIL:cToolTip :="Indique el caracter separador"
  @ oVen:oCESEPMIL:nTop-08,oVen:oCESEPMIL:nLeft SAY "Separador Miles" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 9,1 GET oVen:oCESEPDEC VAR oVen:cESepDec WHEN !oVen:lEDetSep
  oVen:oCESEPDEC:cMsg     :="Indique el caracter separador"
  oVen:oCESEPDEC:cToolTip :="Indique el caracter separador"
  @ oVen:oCESEPDEC:nTop-08,oVen:oCESEPDEC:nLeft SAY "Separador Decimales" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 

  @ 10,1 GET oVen:oCESEPLIS VAR oVen:cESepLis WHEN !oVen:lEDetSep
  oVen:oCESEPLIS:cMsg     :="Indique el caracter separador"
  oVen:oCESEPLIS:cToolTip :="Indique el caracter separador"
  @ oVen:oCESEPLIS:nTop-08,oVen:oCESEPLIS:nLeft SAY "Separador Lista" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 


  SETFOLDER(0)

/*  
  @09, 33 SBUTTON oBtn ;
          SIZE 42, 23 FONT oFont;
          FILE "BITMAPS\RUN.BMP" ;
          LEFT PROMPT "Ejecutar";
          NOBORDER;
          COLORS CLR_BLACK, { CLR_WHITE, CLR_HGRAY, 1 };
          ACTION (CursorWait(),;
                  oVen:dDesde:=CTOD("01/"+STRZERO(oVen:oMes:nAt,2)+"/"+STRZERO(oVen:nAno)),;
                  oVen:dHasta:=FCHFINMES(oVen:dDesde),;
                  oVen:LIBVENTA(oVen),;
                  EJECUTAR("IVALOAD",oVen:dFecha))

  @10, 20 SBUTTON oBtn ;
          SIZE 42, 23 FONT oFont;
          FILE "BITMAPS\XSALIR.BMP" ;
          LEFT PROMPT "Cerrar";
          NOBORDER;
          COLORS CLR_BLACK, { CLR_WHITE, CLR_HGRAY, 1 };
          ACTION (CursorWait(),;
                  oVen:Close())
*/
  oVen:Activate({||oVen:INICIO()})

Return nil

/*
// Grabar Empresa
*/
FUNCTION LibVenta(oVen)
  LOCAL oData,nAt

*  CursorWait()
 
  oData:=DATASET("CSLIBROS","ALL")

  oData:Set("lVControl" ,oVen:lVControl)
  oData:Set("lVSombCab" ,oVen:lVSombCab)
  oData:Set("lVSombTot" ,oVen:lVSombTot)
  oData:Set("nVLenRs"   ,oVen:nVLenRs)
  oData:Set("nVPorAju"  ,oVen:nVPorAju)
  oData:Set("lVSepAju"  ,oVen:lVSepAju)
  oData:Set("lVSepMil"  ,oVen:lVSepMil)
  oData:Set("lVRetAju"  ,oVen:lVRetAju)
  oData:Set("lVRetFec"  ,oVen:lVRetFec) // DR20110915a
  oData:Set("lVRifCab"  ,oVen:lVRifCab)
  oData:Set("lVNumPag"  ,oVen:lVNumPag)
  oData:Set("lVConDin"  ,oVen:lVConDin)
  oData:Set("nVForm"    ,oVen:nVForm)
  oData:Set("lVColRti"  ,oVen:lVColRti)
  oData:Set("lVLibRti"  ,oVen:lVLibRti)

  oData:Set("lEDetIdio"  ,oVen:lEDetIdio)
  oData:Set("lEDetSep"   ,oVen:lEDetSep)
  oData:Set("cEFunCel"   ,oVen:cEFunCel)
  oData:Set("cEFunSum"   ,oVen:cEFunSum)
  oData:Set("cEFunSi"    ,oVen:cEFunSi)
  oData:Set("cESepMil"   ,oVen:cESepMil)
  oData:Set("cESepDec"   ,oVen:cESepDec)
  oData:Set("cESepLis"   ,oVen:cESepLis)

  oData:Save()
  oData:End(.t.)

  EJECUTAR("CSLIBROSVEN3001",oVen) // Todo se unifica en un solo proceso

RETURN .T.

FUNCTION FINDCODSUC()
  oVen:oSucNombre:Refresh(.T.) 

  IF !oVen:cCodSuc==SQLGET("DPSUCURSAL","SUC_CODIGO","SUC_CODIGO"+GetWhere("=",oVen:cCodSuc))
    EVAL(oVen:oCodSuc:bAction)
    RETURN .F.
  ENDIF
RETURN .T.

FUNCTION ChangeControls()
IF oVen:nVForm=2
   oVen:lVSepAju:=.F.
   oVen:oVSepAju:Refresh()
   oVen:lVControl:=.F.
   oVen:oVControl:Refresh()
   oVen:lVLibRti:=.F.
   oVen:oVLibRti:Refresh()
ENDIF
IF !oVen:lVSepAju .OR. oVen:lVLibRti
   oVen:lVRetAju:=.F.
   oVen:oVRetAju:Refresh()
ENDIF
IF oVen:lVRetAju
   oVen:lVLibRti:=.F.
   oVen:oVLibRti:Refresh()
ENDIF
IF oVen:lEDetIdio
   oVen:cEFunCel:=SPACE(10)
   oVen:oCEFUNCEL:Refresh()
   oVen:cEFunSum:=SPACE(10)
   oVen:oCEFUNSUM:Refresh()
   oVen:cEFunSi:=SPACE(10)
   oVen:oCEFUNSi:Refresh()
ENDIF
IF oVen:lEDetSep
   oVen:cESepMil:=" "
   oVen:oCESEPMIL:Refresh()
   oVen:cESepDec:=" "
   oVen:oCESEPDEC:Refresh()
   oVen:cESepLis:=" "
   oVen:oCESEPLIS:Refresh()
ENDIF

RETURN .T.

FUNCTION INICIO()
   LOCAL oCursor,oBar,oBtn,oFont,oCol
   LOCAL oDlg:=oVen:oDlg
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
                  oVen:LIBVENTA(oVen),;
                  EJECUTAR("IVALOAD",oVen:dFecha))

//                  oVen:dDesde:=CTOD("01/"+STRZERO(oVen:oMes:nAt,2)+"/"+STRZERO(oVen:nAno)),;
//                  oVen:dHasta:=FCHFINMES(oVen:dDesde),;


   oBtn:cToolTip:="Guardar"

   oVen:oBtnRun:=oBtn


   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\XSALIR.BMP";
          TOP PROMPT "Cerrar"; 
          ACTION (oVen:Cancel()) CANCEL

   oBar:SetColor(CLR_BLACK,oDp:nGris)

   AEVAL(oBar:aControls,{|o,n| o:SetColor(CLR_BLACK,oDp:nGris) })

RETURN .T.



