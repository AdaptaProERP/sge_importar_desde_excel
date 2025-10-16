// Programa   : 
// Fecha/Hora : 05/12/2007 14:54:06
// Propósito  : 
// Aplicación : 
// Tipo       : STD00000
// Creado Por : 
// Observación:
#include "dpxbase.ch"

PROCE MAIN(oVen)
   LOCAL oFont,oFontT,oFontB,oFontD,aAlias:={},lControl:=.T.
   LOCAL nCantLines:=0,nCantCols:=0,nColStep:=0,nRowStep:=0,nLinea:=0,nTc:=0,nPage:=0,nRow:=0,nCol:=0,nCounter:=0
   LOCAL cBmp:="",cEstatus,oPen,oShadow,aCabecera,aRangos,cCab,aCab,aCab1,aCab2,lHubo:=.F.,lPrimera:=.T.,aMatriz,nWfont,nHfont,nBaseRed
   LOCAL aPie,aCondicion,nTotal,nMonto,nWid,nHei,nlin2,cAliFac,cAliRet,cAliCli,cAliMem,nI,nJ,nColIni,nColFin,aDet,nX,cImp,nFactor,nPrimCol:=6,nZoom:=100,lEspanol,lMiles:=.T.,nSalida
   LOCAL cRif,cRsocial,cPlanilla,cCadena1,cCadena2,cCadena3,cCadena4,cDocFis,cAfec,nTotDoc,nExento,nBase,nAlic,nIva,nIvaRet,nIvaPer,nDescue,nRecarg,nLenrs,cDocNum,cLetra1,cLetra2,cLetra3,cLetra4
   LOCAL nTotal1,nTotal2,nTotal3,nTotal4,nTotal5,nTotal6
   LOCAL nInternaNG,nExporta,nTotGen,nTotGenAdi,nTotReduc,nIvaGen,nIvaGenAdi,nIvaReduc,nRetGen,nRetGenAdi,nRetReduc,nPerGen,nPerGenAdi,nPerReduc,lSepAju:=.F.,lRetAju:=.F.,lNumera:=.F.
   LOCAL aLim,nCol1,nCol2,nCol3,nCol4,nCol5,nCol6,nCol7,nCol8,nCol9
   LOCAL aLim2,nCol21,nCol22,nCol23
   LOCAL nDemoMax:=2,nDemoLin:=1,cLetraBase,nBaseCol
   LOCAL oExcel,oHoja,oHoja2,oHojax,oHojaSuma,oLanguage,nRowIni,nRowIni2,nRowFin,nRowFin2,cRange1,cRange2,cRange3,cRange4,aResumen:=ARRAY(5,4),nRowInicab,lSombra1:=.T.,lSombra2:=.T.,lDinamic:=.T.,lContri
   // Revisión 03/03/2006
   LOCAL lRif:=.T.,cEmpRif:="",nDinamic
   LOCAL cAliIva
   // Internacional por configuración
   LOCAL lDetecIdi:=.T.,lDetecSep:=.T.
   // Compatibilidad
   LOCAL cFunCelda,cFunSuma,cFunSi,cMascPor,cMascMon,cSepFun,cSepMil,cSepDec

   IF !EsDemo .AND. EsDemoVe
      MsgInfo("Su Licencia está limitada. Se emitirá un Libro de Ventas Demostrativo")
   ENDIF

   nSalida:=DySalida("Libro de Ventas. Formato Separado de Documentos",.T.,.T.,.F.)

   IF VALTYPE(nSalida)<>"N" .OR. nSalida=0
      MsgAlert("El Informe ha sido cancelado","Alerta")
      RETURN
   ENDIF

   dFecha1:=IIF( EMPTY(dFecha1),FCHINIMES(FECHADIA) ,dFecha1 )
   dFecha2:=IIF( EMPTY(dFecha2),FCHFINMES(FECHADIA) ,dFecha2 )

   FOR nI := 1 TO LEN(aResumen)
       aResumen[nI,1]:=""
       aResumen[nI,2]:=""
       aResumen[nI,3]:=""
       aResumen[nI,4]:=""
   NEXT

   CNF0001:=.T.
   CNF0002:=.T.
   CNF0003:=.T.
   CNF0006:=40
   CNF0008:=100
   CNF0010:=.F.
   CNF0012:=.F.
   CNF0014:=.T.
   CNF0016:=.T.
   CNF0018:=.T.
   CNF0019:=.T.
   CNF0020:=PADR("CELDA",10)
   CNF0021:=PADR("SUMA",10)
   CNF0022:="."
   CNF0023:=","
   CNF0024:=";"
   CNF0025:=PADR("SI",10)
   CNF0026:=.F.
   CNF0031:=.T.
   IF FILE(cUbicData+"CONFIG.DAT")
      RESTORE FROM (cUbicData+"CONFIG.DAT") ADDITIVE
   ENDIF
   lControl:=CNF0001
   lSombra1:=CNF0002
   lSombra2:=CNF0003
   nZoom:=CNF0008
   lSepAju:=CNF0010
   lNumera:=CNF0012
   lDinamic:=CNF0014
   lMiles:=CNF0016
   lDetecIdi:=CNF0018
   lDetecSep:=CNF0019
   cFunCelda:=ALLTRIM(CNF0020)
   cFunSuma:=ALLTRIM(CNF0021)
   cFunSi:=ALLTRIM(CNF0025)
   cSepMil:=CNF0022
   cSepDec:=CNF0023
   cSepFun:=CNF0024
   nLenRs:=CNF0006
   lRetAju:=IIF(!lSepAju,.F.,CNF0026 )
   lRif:=CNF0031

   nDinamic:=IIF( lRif,10 ,9 )

   IF !FILE(cDpUbic+"DpRetIvc"+DpExt)
      MsgStop("Su versión del Sistema Administrativo no está apta para manejar retenciones IVA en Ventas")
      RETURN .F.
   ENDIF

   cAliFac  := CnwOpenDbf(oWndMain,cDpUbic+"DpFac"+DpExt,1,.T.,.F.,IIF( DpVia="DBFCDX", ,{cDpUbic+"DpFac1",cDpUbic+"DpFac2",cDpUbic+"DpFac3"} ),DpVia)
   cAliCli  := CnwOpenDbf(oWndMain,cDpUbic+"DpCli"+DpExt,1,.T.,.F.,IIF( DpVia="DBFCDX", ,{cDpUbic+"DpCli1",cDpUbic+"DpCli2"} ),DpVia)
   cAliRet  := CnwOpenDbf(oWndMain,cDpUbic+"DpRetIvc"+DpExt,1,.T.,.F.,IIF( DpVia="DBFCDX", ,{cDpUbic+"DpRetIc1",cDpUbic+"DpRetIc2",cDpUbic+"DpRetIc3",cDpUbic+"DpRetIc4"} ),DpVia)
   cAliMem  := CnwOpenDbf(oWndMain,cDpUbic+"DpMemo"+DpExt,1,.T.,.F.,IIF( DpVia="DBFCDX", ,{cDpUbic+"DpMemo1"} ),DpVia)
   cAliIva  := CnwOpenDbf(oWndMain,cDpUbic+"DpEmpIva"+DpExt,1,.T.,.F.,IIF( DpVia="DBFCDX", , ),DpVia)
   AADD(aAlias,cAliFac)
   AADD(aAlias,cAliCli)
   AADD(aAlias,cAliRet)
   AADD(aAlias,cAliMem)
   AADD(aAlias,cAliIva)
   IF EMPTY(cAliFac) .OR. EMPTY(cAliCli) .OR. EMPTY(cAliRet) .OR. EMPTY(cAliMem) .OR. EMPTY(cAliIva)
      MsgStop("No se pudo abrir alguna de las tablas asociadas al reporte","No es posible continuar")
      AEVAL(aAlias,{|x,y| IIF( !EMPTY(aAlias[y]),(aAlias[y])->(DBCLOSEAREA()),NIL)})
      RETURN
   ENDIF

   (cAliFac)->(ORDSETFOCUS(3))
   (cAliCli)->(ORDSETFOCUS(1))
   (cAliRet)->(ORDSETFOCUS(1))
   (cAliMem)->(ORDSETFOCUS(1))

   cEmpRif:=IIF( !(cAliIva)->(EOF()) .AND. !(cAliIva)->(BOF()), (cAliIva)->EMP_RIF,SPACE(15) )

   // Comentarios de comandos validos
/*
   msgstop(oExcel:WorkSheets(1))
   msgstop(oExcel:WorkSheets(1):name)
   oHoja2:=oExcel:Sheets:Add()
   oHoja := oExcel:Get( "ActiveSheet" )
   msgstop(oExcel:Name)
 */
   oExcel := TOleAuto():New( "Excel.Application" )
   IF lDetecIdi
      oLanguage:=oExcel:Get("LanguageSettings")
      lEspanol:=IsExcelSpa(oLanguage:LanguageID(1))
   ENDIF
   IF lDetecSep
      cSepDec:=oExcel:Get("International",3)
      cSepMil:=oExcel:Get("International",4)
      cSepFun:=oExcel:Get("International",5)
   ENDIF
   cMascPor:="##0"+cSepDec+"00%"
   cMascMon:="###"+cSepMil+"###"+cSepMil+"##0"+cSepDec+"00"
   IF !lMiles
      cMascMon:="########0"+cSepDec+"00"
   ENDIF

   IF lDetecIdi
      cFunCelda:="CELDA"
      cFunSuma:="SUMA"
      cFunSi:="SI"
      IF !lEspanol
         cFunCelda:="CELL"
         cFunSuma:="SUM"
         cFunSi:="IF"
      ENDIF
   ENDIF

   oExcel:WorkBooks:Add()
   oHojaSuma:=oExcel:WorkSheets:Add()
   oHojaSuma:Name:="Valores"
   IF lSepAju
      oHoja2:=oExcel:WorkSheets:Add()
      oHoja2:Name:="Ajustes"
   ENDIF
   oHoja:=oExcel:WorkSheets:Add()
   oHoja:Name:="Ventas"
   oExcel:Visible := .T.
   oHoja:Cells:Font:Name := "Arial"
   oHoja:Cells:Font:Size := 8
   oHoja:PageSetup:Set("Zoom",nZoom)
   oHoja:PageSetup:Set("Orientation",2 )
   oHoja:PageSetup:Set("PaperSize" , 1 )
   oHoja:PageSetup:Set("PrintTitleRows","A1:A"+ALLTRIM(STR(nDinamic)))
   IF lNumera
      oHoja:PageSetup:Set("RightHeader", "&P" )
   ENDIF

   nLinea:=0
   nTc:=0
   nPage:=0
   nRow:=0
   nCol:=0
   nCounter:=0

   nTotal:=0

   aRangos:={}
   IF !EMPTY(dFecha1)
      AADD(aRangos,{"Fecha Desde: "+DTOC(dFecha1)+" Hasta: "+DTOC(dFecha2),oFont})
   ENDIF

   aCab1:={}
   AADD(aCab1,{IIF( EsDemo .OR. EsDemoVe,"Empresa Demostrativa, C.A." , ALLTRIM(cEmpresa)),oFontT})
   IF lRif
      AADD(aCab1,{IIF( EsDemo .OR. EsDemoVe,"J-99999999999-9" , ALLTRIM(cEmpRif)),oFontT})
   ENDIF
   AADD(aCab1,{"Libro de Ventas"+IIF( EsDemo .OR. EsDemoVe,"-Versión Demostrativa" ,"" ),oFontT})

   aCabecera:={}
   //Primera Linea
   AADD(aCabecera,space(7))
   AADD(aCabecera,padc('Fecha',11))
   AADD(aCabecera,space(15))
   AADD(aCabecera,space(nLenRs))
   AADD(aCabecera,space(1))
   AADD(aCabecera,Padc('Núm.Planilla',15))
   AADD(aCabecera,space(1))
   AADD(aCabecera,padc("Número",10))
   AADD(aCabecera,space(10))
   AADD(acabecera,padc("Número",10))
   AADD(acabecera,space(10))
   AADD(acabecera,space(10))
   AADD(aCabecera,padc("Número de",10))
   AADD(aCabecera,padc("Total Ventas",15))
   AADD(aCabecera,padc("Ventas",15))
   AADD(aCabecera,"")
   AADD(aCabecera,"")
   AADD(aCabecera,"")
   AADD(aCabecera,space(7))
   AADD(aCabecera,space(15))
   AADD(aCabecera,padc("IVA Retenido",15))
   AADD(aCabecera,space(15))
   //Segunda Linea
   AADD(aCabecera,padc("Oper",7))
   AADD(aCabecera,padc('de la',11))
   AADD(aCabecera,space(15))
   AADD(aCabecera,space(nLenRs))
   AADD(aCabecera,space(1))
   AADD(aCabecera,padc("de Exportación",15))
   AADD(aCabecera,space(1))
   AADD(aCabecera,padc("de",10))
   AADD(aCabecera,padc("Núm.Ctrol.",10))
   AADD(aCabecera,padc("Nota Déb.",10))
   AADD(aCabecera,padc("Número de",10))
   AADD(aCabecera,padc("Tipo de",10))
   AADD(aCabecera,padc("Factura",10))
   AADD(aCabecera,padc("Incluyendo",15))
   AADD(aCabecera,padc("Internas",15))
   AADD(aCabecera,"Base Imponible")
   AADD(aCabecera,"")
   AADD(aCabecera,"")
   AADD(aCabecera,padc("%",7))
   AADD(aCabecera,padc("Impuesto",15))
   AADD(aCabecera,padc("Por el",15))
   AADD(aCabecera,padc("IVA",15))
   //Tercera Linea
   AADD(aCabecera,padc("Nro.",7))
   AADD(aCabecera,padc('Factura',11))
   AADD(aCabecera,padc("R.I.F.",15))
   AADD(aCabecera,padc('Nombre o Razón Social',nLenRs))
   AADD(aCabecera,space(1))
   AADD(aCabecera,padc("["+"Forma D"+"]",15))
   AADD(aCabecera,space(1))
   AADD(aCabecera,padc("Factura",10))
   AADD(aCabecera,padc("de Factura",10))
   AADD(aCabecera,SPACE(10))
   AADD(aCabecera,padc("Nota Crdto",10))
   AADD(aCabecera,padc("Transacc.",10))
   AADD(aCabecera,padc("Afectada",10))
   AADD(acabecera,padc("el IVA",15))
   AADD(aCabecera,padc("No Gravadas",15))
   AADD(aCabecera,"Exportaciones")
   AADD(aCabecera,"Contribuyentes")
   AADD(aCabecera,"No Contribuyentes")
   AADD(aCabecera,padc("Alicuota",7))
   AADD(aCabecera,padc("IVA",15))
   AADD(aCabecera,padc("Comprador",15))
   AADD(aCabecera,padc("Percibido",15))
   FOR nI := 1 TO LEN(aCabecera)
       aCabecera[nI]:=ALLTRIM(aCabecera[nI])
   NEXT
   nRow:=1
   FOR nI := 1 TO LEN(aCab1)
       oHoja:Cells( nRow, 1):Value := aCab1[nI,1]
       oHoja:Cells( nRow, 1 ):Font:Size := IIF( nI=1,12 ,10 )
       nRow++
   NEXT
   FOR nI := 1 TO LEN(aRangos)
       oHoja:Cells( nRow, 1):Value := aRangos[nI,1]
       oHoja:Cells( nRow, 1 ):Font:Size := 8
       oHoja:Cells( nRow, 1 ):Font:Bold := .T.
       nRow++
   NEXT
   nRow++
   cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":T"+ALLTRIM(STR(INT(nRow+1)))
   oHoja:Range( cRange1 ):Borders():LineStyle  := 1
   oHoja:Cells(nRow,16):Value:="VENTAS INTERNAS O"
   oHoja:Range( "P"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
   oHoja:Cells( nRow, 16):Set("HorizontalAlignment",3)
   IF lSombra1
      oHoja:Range( cRange1 ):Interior:Color := 12632256
   ENDIF
   nRow++
   oHoja:Cells(nRow,16):Value:="EXPORTACIONES GRAVADAS"
   oHoja:Range( "P"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
   oHoja:Cells( nRow, 16):Set("HorizontalAlignment",3)
   cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":T"+ALLTRIM(STR(INT(nRow+1)))
   IF lSombra1
      oHoja:Range( cRange1 ):Interior:Color := 12632256
   ENDIF
   nRow++
   nRowIni:=nRow
   nRowIniCab:=nRow
   nJ:=1
   FOR nI := 1 TO LEN(aCabecera)
       IF nI=16 .OR. nI=32
          oHoja:Range( "P"+ALLTRIM(STR(nRow))+":R"+ALLTRIM(STR(nRow))):Merge()
          oHoja:Cells( nRow, 16):Set("HorizontalAlignment",3)
       ENDIF
       oHoja:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/3)) ):Value := aCabecera[nI]
       oHoja:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/3)) ):Set("HorizontalAlignment",3)
       IF LEN(aCabecera)/3=nI/nJ
          nRow++
          nJ++
       ENDIF
   NEXT
   cRange1:="A"+ALLTRIM(STR(INT(nRowIni)))+":V"+ALLTRIM(STR(INT(nRowIni+2)))
   oHoja:Range( cRange1 ):Borders():LineStyle  := 1
   IF lSombra1
      oHoja:Range( cRange1 ):Interior:Color := 12632256
   ENDIF
   nRowIni:=nRow
   nRowFin:=nRow

   aPie:={}

   lHubo:=.F.

   nI:=1
   nTotal1:=0
   nTotal2:=0
   nTotal3:=0
   nTotal4:=0
   nTotal5:=0
   nTotal6:=0

   nInternaNG:=0
   nExporta:=0
   nTotGen:=0
   nTotGenAdi:=0
   nTotReduc:=0
   nIvaGen:=0
   nIvaGenAdi:=0
   nIvaReduc:=0
   nRetGen:=0
   nRetGenAdi:=0
   nRetReduc:=0
   nPerGen:=0
   nPerGenAdi:=0
   nPerReduc:=0
   lPrimera:=.T.
   (cAliFac)->(DBSEEK(DTOS(dFecha1),.T.))
   DO WHILE ((!EsDemo .AND. !EsDemoVe) .OR. nDemoLin<=nDemoMax) .AND. !(cAliFac)->(EOF()) .AND. (cAliFac)->FAC_FECHA<=dFecha2
      IF (cAliFac)->FAC_FECHA<dFecha1
         (cAliFac)->(DBSKIP())
         LOOP
      ENDIF
      IF !(cAliFac)->FAC_TIPO$"FC,DB,CR"
         (cAliFac)->(DBSKIP())
         LOOP
      ENDIF
      IF !EMPTY(cNumero1) .AND. ((cAliFac)->FAC_NUMERO<cNumero1 .OR. (cAliFac)->FAC_NUMERO>cNumero2)
         (cAliFac)->(DBSKIP())
         LOOP
      ENDIF

      cDocNum:=(cAliFac)->FAC_NUMERO

      nFactor:=1
      IF (cAliFac)->FAC_TIPO="CR"
         nFactor:=-1
      ENDIF

      lPrimera:=.F.
      cRif:=SPACE(15)
      cRsocial:=SPACE(nLenRs)
      cPlanilla:=IIF( (cAliFac)->(TYPE("FAC_NUMPLA"))="C",PADR((cAliFac)->FAC_NUMPLA,15) ,SPACE(15) )
      cDocFis:=PADR(IIF( !EMPTY((cAliFac)->FAC_DOCFIS).OR.!lControl,(cAliFac)->FAC_DOCFIS,(cAliFac)->FAC_NUMERO),10)
      cAfec:=PADR(IIF( (cAliFac)->FAC_TIPO$"CR,DB",(cAliFac)->FAC_NOTAEN ," " ),10)
      nTotDoc:=(cAliFac)->FAC_NETO
      nDescue:=((cAliFac)->FAC_DESCUE*(cAliFac)->FAC_BRUTO/100)
      nRecarg:=((cAliFac)->FAC_RECARG*(cAliFac)->FAC_BRUTO/100)
      nBase:=(cAliFac)->FAC_BRUTO-nDescue+nRecarg+(cAliFac)->FAC_OTROS-(cAliFac)->FAC_EXONER
      nExento:=(cAliFac)->FAC_EXONER
      nIva:=(cAliFac)->FAC_IMPTO+(cAliFac)->FAC_IMPICS+(cAliFac)->FAC_ICS_A2+(cAliFac)->FAC_IVARED
      nAlic:=RATA(nIva,nBase)
      nIvaRet:=0
      lContri:= .F.
      IF (cAliCli)->(DBSEEK((cAliFac)->FAC_CODCLI)) .AND. (cAliCli)->(TYPE("CLI_COLLIB"))="C"
         lContri:=( (cAliCli)->CLI_COLLIB<>"N" )
      ELSE
         lContri:=( (cAliFac)->FAC_CONTRI<>"N" )
      ENDIF
      IF LEFT((cAliFac)->FAC_CODCLI,1)="0" .AND. LEN(ALLTRIM((cAliFac)->FAC_CODCLI))=1 .AND. (cAliFac)->(TYPE("FAC_COLLIB"))="C"
         lContri:=( (cAliFac)->FAC_COLLIB<>"N" )
      ENDIF
      cLetraBase:=IIF( !lContri,"R" ,"Q" )
      nBaseCol:=IIF( !lContri,18 ,17 )
      IF (cAliFac)->FAC_ESTADO<>"A"
         nTotGen+=IIF( (cAliFac)->FAC_IMPTO>0 .AND. (cAliFac)->FAC_IMPICS=0 .AND. (cAliFac)->FAC_ICS_A2=0,(nBase*nFactor) , 0)
         nIvaGen+=IIF( (cAliFac)->FAC_IMPTO>0 .AND. (cAliFac)->FAC_IMPICS=0 .AND. (cAliFac)->FAC_ICS_A2=0,(((cAliFac)->FAC_IMPTO+(cAliFac)->FAC_IVARED)*nFactor) , 0)
         nTotGenAdi+=IIF( (cAliFac)->FAC_IMPTO>0 .AND. ((cAliFac)->FAC_IMPICS>0 .OR. (cAliFac)->FAC_ICS_A2>0),(nBase*nFactor) , 0)
         nIvaGenAdi+=IIF( (cAliFac)->FAC_IMPTO>0 .AND. ((cAliFac)->FAC_IMPICS>0 .OR. (cAliFac)->FAC_ICS_A2>0),(((cAliFac)->FAC_IMPTO+(cAliFac)->FAC_IMPICS+(cAliFac)->FAC_ICS_A2)*nFactor) , 0)
         nTotReduc+=IIF( (cAliFac)->FAC_IVARED>0 .AND. (cAliFac)->FAC_IMPTO<=0,(nBase*nFactor) , 0)
         nIvaReduc+=IIF( (cAliFac)->FAC_IVARED>0 .AND. (cAliFac)->FAC_IMPTO<=0,((cAliFac)->FAC_IVARED*nFactor) , 0)
      ENDIF

      IF (cAliRet)->(DBSEEK((cAliFac)->FAC_TIPO+(cAliFac)->FAC_CODCLI+(cAliFac)->FAC_NUMERO))
         cPlanilla:=PADR((cAliRet)->RTI_NUMEXP,15)
         cDocFis:=PADR((cAliRet)->RTI_DOCFIS,10)
         cAfec:=PADR((cAliRet)->RTI_FACASO,10)
         *nTotDoc:=(cAliRet)->RTI_MTODOC
         *nBase:=(cAliRet)->RTI_BASE
         *nExento:=(cAliRet)->RTI_EXENTO
         *nAlic:=(cAliRet)->RTI_PORIVA
         *nIva:=(cAliRet)->RTI_MTOIVA
         // MODIFICACION RETENCIONES NO DECLARADAS EN SU MOMENTO
         IF !lRetAju .OR. LEFT(DTOS((cAliFac)->FAC_FECHA),6)>=LEFT(DTOS((cAliRet)->RTI_FECHA),6)
         // FIN MODIFICACION
            nIvaRet:=(cAliRet)->RTI_MONTO
            IF (cAliFac)->FAC_ESTADO<>"A"
               nRetGen+=IIF( (cAliFac)->FAC_IMPTO>0 .AND. (cAliFac)->FAC_IMPICS=0 .AND. (cAliFac)->FAC_ICS_A2=0,((cAliRet)->RTI_MONTO*nFactor) , 0)
               nRetGenAdi+=IIF( (cAliFac)->FAC_IMPTO>0 .AND. ((cAliFac)->FAC_IMPICS>0 .OR. (cAliFac)->FAC_ICS_A2>0),((cAliRet)->RTI_MONTO*nFactor) , 0)
               nRetReduc+=IIF( (cAliFac)->FAC_IVARED>0,((cAliRet)->RTI_MONTO*nFactor) , 0)
            ENDIF
         ENDIF
      ENDIF

      IF (cAliCli)->(DBSEEK((cAliFac)->FAC_CODCLI))
         cRif:=PADR((cAliCli)->CLI_RIF,15)
         cRsocial:=OemToAnsi(PADR((cAliCli)->CLI_NOMBRE,nLenRs))
         IF (cAliCli)->CLI_ENOTRA="E" .AND. nIva=0
            *cDocNum:=IIF( (cAliFac)->FAC_TIPO="FC",PADR("N/A",8),cDocNum)
            *cDocFis:=IIF( (cAliFac)->FAC_TIPO="FC",PADR("N/A",10),cDocNum)
            nExento:=0
            nBase:=(cAliFac)->FAC_NETO
            cLetraBase:="P"
            nBaseCol:=16
            nExporta+=((cAliFac)->FAC_NETO*nFactor)
         ENDIF
      ENDIF

      IF (cAliFac)->FAC_NUMMEM>0 .AND. LEFT((cAliFac)->FAC_CODCLI,1)="0" .AND. LEN(ALLTRIM((cAliFac)->FAC_CODCLI))=1 .AND. (cAliMem)->(DBSEEK((cAliFac)->FAC_NUMMEM))
         cRsocial:=OemToAnsi(PADR((cAliMem)->MEM_DESCRI,nLenRs))
         cRif:=PADR(MEMOLINE((cAliMem)->MEM_MEMO,,5),15) // PADR((cAliCli)->CLI_RIF,15)
      ENDIF

      nTotDoc:=nTotDoc*nFactor
      nBase:=nBase*nFactor
      nExento:=nExento*nFactor
      nIva:=nIva*nFactor
      nIvaRet:=nIvaRet*nFactor

      nIvaPer:=nIva-nIvaRet

      IF (cAliFac)->FAC_ESTADO<>"A" .AND. (cAliFac)->FAC_IMPTO>0 .AND. (cAliFac)->FAC_IMPICS=0 .AND. (cAliFac)->FAC_IMPICS=0 .AND. (cAliFac)->FAC_ICS_A2=0
         aResumen[3,1]:=IIF( EMPTY(aResumen[3,1]),"=" ,aResumen[3,1])
         aResumen[3,2]:=IIF( EMPTY(aResumen[3,2]),"=" ,aResumen[3,2])
         aResumen[3,3]:=IIF( EMPTY(aResumen[3,3]),"=" ,aResumen[3,3])
         aResumen[3,4]:=IIF( EMPTY(aResumen[3,4]),"=" ,aResumen[3,4])
         aResumen[3,1]:=aResumen[3,1]+"+"+cLetraBase+ALLTRIM(STR(INT(nRow)))
         aResumen[3,2]:=aResumen[3,2]+"+T"+ALLTRIM(STR(INT(nRow)))
         aResumen[3,3]:=aResumen[3,3]+"+U"+ALLTRIM(STR(INT(nRow)))
         aResumen[3,4]:=aResumen[3,4]+"+V"+ALLTRIM(STR(INT(nRow)))
         oHojaSuma:Cells( nRow, 01):Value := "VIAG"
      ENDIF

      IF (cAliFac)->FAC_ESTADO<>"A" .AND. (cAliFac)->FAC_IMPTO>0 .AND. ((cAliFac)->FAC_IMPICS>0 .OR. (cAliFac)->FAC_ICS_A2>0)
         aResumen[4,1]:=IIF( EMPTY(aResumen[4,1]),"=" ,aResumen[4,1])
         aResumen[4,2]:=IIF( EMPTY(aResumen[4,2]),"=" ,aResumen[4,2])
         aResumen[4,3]:=IIF( EMPTY(aResumen[4,3]),"=" ,aResumen[4,3])
         aResumen[4,4]:=IIF( EMPTY(aResumen[4,4]),"=" ,aResumen[4,4])
         aResumen[4,1]:=aResumen[4,1]+"+"+cLetraBase+ALLTRIM(STR(INT(nRow)))
         aResumen[4,2]:=aResumen[4,2]+"+T"+ALLTRIM(STR(INT(nRow)))
         aResumen[4,3]:=aResumen[4,3]+"+U"+ALLTRIM(STR(INT(nRow)))
         aResumen[4,4]:=aResumen[4,4]+"+V"+ALLTRIM(STR(INT(nRow)))
         oHojaSuma:Cells( nRow, 01):Value := "VIGA"
      ENDIF

      IF (cAliFac)->FAC_ESTADO<>"A" .AND. (cAliFac)->FAC_IVARED>0 .AND. (cAliFac)->FAC_IMPTO<=0
         aResumen[5,1]:=IIF( EMPTY(aResumen[5,1]),"=" ,aResumen[5,1])
         aResumen[5,2]:=IIF( EMPTY(aResumen[5,2]),"=" ,aResumen[5,2])
         aResumen[5,3]:=IIF( EMPTY(aResumen[5,3]),"=" ,aResumen[5,3])
         aResumen[5,4]:=IIF( EMPTY(aResumen[5,4]),"=" ,aResumen[5,4])
         aResumen[5,1]:=aResumen[5,1]+"+"+cLetraBase+ALLTRIM(STR(INT(nRow)))
         aResumen[5,2]:=aResumen[5,2]+"+R"+ALLTRIM(STR(INT(nRow)))
         aResumen[5,3]:=aResumen[5,3]+"+S"+ALLTRIM(STR(INT(nRow)))
         aResumen[5,4]:=aResumen[5,4]+"+T"+ALLTRIM(STR(INT(nRow)))
         oHojaSuma:Cells( nRow, 01):Value := "VIAR"
      ENDIF

      IF (cAliFac)->FAC_ESTADO="A"
         IF (cAliFac)->FAC_TIPO="FC"
            cAfec:=PADR((cAliFac)->FAC_NUMERO,10)
         ENDIF
         nTotDoc:=0
         nBase:=0
         nExento:=0
         nIva:=0
         nAlic:=0
         nIvaRet:=0
         nIvaPer:=0
      ENDIF

      cCadena1:=SPACE(54)
      IF (cAliFac)->FAC_TIPO="FC"
         cCadena1:=PADC(cDocNum,10)+" "+PADC(cDocFis,10)+" "+SPACE(10)+" "+SPACE(10)+" "+PADC(IIF( A->FAC_ESTADO="A","03-anu" ,"01-reg" ),10)
      ENDIF
      IF (cAliFac)->FAC_TIPO="DB"
         cCadena1:=SPACE(10)+" "+PADC(cDocFis,10)+" "+PADC(cDocNum,10)+" "+SPACE(10)+" "+PADC(IIF( A->FAC_ESTADO="A","03-anu" ,"01-reg" ),10)
      ENDIF
      IF (cAliFac)->FAC_TIPO="CR"
         cCadena1:=SPACE(10)+" "+PADC(cDocFis,10)+" "+SPACE(10)+" "+PADC(cDocNum,10)+" "+PADC(IIF( A->FAC_ESTADO="A","03-anu" ,"01-reg" ),10)
      ENDIF
      cCadena2:=cAfec+" "+PADL(ALLTRIM(TRANSFORM(nTotDoc,"999,999,999.99")),14)
      cCadena3:=PADL(ALLTRIM(TRANSFORM(nExento,"999,999,999.99")),14)+" "+PADL(ALLTRIM(TRANSFORM(nBase,"999,999,999.99")),14)
      cCadena4:=PADL(ALLTRIM(TRANSFORM(nIvaRet,"999,999,999.99")),14)+" "+PADL(ALLTRIM(TRANSFORM(nIvaPer,"999,999,999.99")),14)
      aDet:={}
      oHoja:Cells( nRow, 01):Value := IIF( lDinamic,'='+cFunCelda+'("ROW"'+cSepFun+'A'+ALLTRIM(STR(nRow))+')-'+ALLTRIM(STR(nDinamic)),nI)
      oHoja:Cells( nRow, 02):Value := (cAliFac)->FAC_FECHA
      oHoja:Cells( nRow, 03):Value := ALLTRIM(cRif)
      oHoja:Cells( nRow, 04):Value := ALLTRIM(cRSocial)
      oHoja:Cells( nRow, 05):Value := " "
      oHoja:Cells( nRow, 06):Value := ALLTRIM(cPlanilla)
      oHoja:Cells( nRow, 07):Value := " "
      oHoja:Cells( nRow, 08):Set( "NumberFormat", "@")
      oHoja:Cells( nRow, 08):Value := ALLTRIM(LEFT(cCadena1,10))
      oHoja:Cells( nRow, 09):Set( "NumberFormat", "@")
      oHoja:Cells( nRow, 09):Value := ALLTRIM(SUBS(cCadena1,12,10))
      oHoja:Cells( nRow, 10):Set( "NumberFormat", "@")
      oHoja:Cells( nRow, 10):Value := ALLTRIM(SUBS(cCadena1,23,10))
      oHoja:Cells( nRow, 11):Set( "NumberFormat", "@")
      oHoja:Cells( nRow, 11):Value := ALLTRIM(SUBS(cCadena1,34,10))
      oHoja:Cells( nRow, 12):Set( "NumberFormat", "@")
      oHoja:Cells( nRow, 12):Value := ALLTRIM(RIGHT(cCadena1,10))
      oHoja:Cells( nRow, 13):Set( "NumberFormat", "@")
      oHoja:Cells( nRow, 13):Value := ALLTRIM(LEFT(cCadena2,10))
      oHoja:Cells( nRow, 14):Value := nTotDoc
      oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
      oHoja:Cells( nRow, 15):Value := nExento
      oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
      oHoja:Cells( nRow, nBaseCol):Value := nBase
      oHoja:Cells( nRow, nBaseCol):Set( "NumberFormat", cMascMon )
      oHoja:Cells( nRow, 19):Value := nAlic/100
      oHoja:Cells( nRow, 19):Set( "NumberFormat", cMascPor )
      oHoja:Cells( nRow, 20):Value := nIva
      oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
      oHoja:Cells( nRow, 21):Value := nIvaRet
      oHoja:Cells( nRow, 21):Set( "NumberFormat", cMascMon )
      oHoja:Cells( nRow, 22):Value := nIvaPer
      oHoja:Cells( nRow, 22):Set( "NumberFormat", cMascMon )
      nRowFin:=nRow
      cRange1:="A"+ALLTRIM(STR(INT(nRow)))+":V"+ALLTRIM(STR(INT(nRow)))
      oHoja:Range( cRange1 ):Borders():LineStyle  := 1
      // VIAG, VIGA o VIRE
      // nBase
      // nIva
      // nIvaRet
      // nIvaPer
      oHojaSuma:Cells( nRow, 02):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+cSepFun+oHoja:Name+'!'+cLetraBase+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 03):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+cSepFun+oHoja:Name+'!'+cLetraBase+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 04):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+cSepFun+oHoja:Name+'!'+cLetraBase+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 05):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+cSepFun+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 06):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+cSepFun+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 07):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+cSepFun+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 08):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+cSepFun+oHoja:Name+'!U'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 09):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+cSepFun+oHoja:Name+'!U'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 10):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+cSepFun+oHoja:Name+'!U'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 11):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+cSepFun+oHoja:Name+'!V'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 12):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+cSepFun+oHoja:Name+'!V'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      oHojaSuma:Cells( nRow, 13):Value := '=+'+cFunSI+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+cSepFun+oHoja:Name+'!V'+ALLTRIM(STR(nRow))+cSepFun+'0)'
      nColIni:=nColStep*nPrimCol
      nTotal1+=nTotDoc
      nTotal2+=nExento
      nTotal3+=nBase
      nTotal4+=nIva
      nTotal5+=nIvaRet
      nTotal6+=nIvaPer
      nI++
      nDemoLin++
      nRow++
      (cAliFac)->(DBSKIP())
   ENDDO
   cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
   aResumen[1,1]:=IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   aResumen[2,1]:=nExporta
   nInternaNG:=nTotal2
   nPerGen:=nIvaGen-nRetGen
   nPerGenAdi:=nIvaGenAdi-nRetGenAdi
   nPerReduc:=nIvaReduc-nRetReduc
   aDet:={}
   cRange1:="M"+ALLTRIM(STR(INT(nRow)))+":V"+ALLTRIM(STR(INT(nRow)))
   oHoja:Range( cRange1 ):Borders():LineStyle  := 1
   oHoja:Cells( nRow, 13):Value := "Totales"
   oHoja:Cells( nRow, 13 ):Font:Bold := .T.
   cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":N"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 14):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
   cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 15):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
   cRange2:="P"+ALLTRIM(STR(INT(nRowIni)))+":P"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 16):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   oHoja:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
   cRange2:="Q"+ALLTRIM(STR(INT(nRowIni)))+":Q"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 17):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   oHoja:Cells( nRow, 17):Set( "NumberFormat", cMascMon )
   cRange2:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 18):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
   cRange2:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 20):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
   cRange2:="U"+ALLTRIM(STR(INT(nRowIni)))+":U"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 21):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   oHoja:Cells( nRow, 21):Set( "NumberFormat", cMascMon )
   cRange2:="V"+ALLTRIM(STR(INT(nRowIni)))+":V"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 22):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
   oHoja:Cells( nRow, 22):Set( "NumberFormat", cMascMon )
   cRange2:="M"+ALLTRIM(STR(INT(nRow)))+":V"+ALLTRIM(STR(INT(nRow)))
   IF lSombra2
      oHoja:Range( cRange2 ):Interior:Color := 12632256
   ENDIF
   nRow+=2
   oHoja:Cells( nRow, 18):Value :="IVA Retenido por"
   oHoja:Cells( nRow, 18):Set("HorizontalAlignment",3)
   oHoja:Cells( nRow, 18):Font:Bold := .T.
   oHoja:Cells( nRow, 20):Value :="IVA de Ventas"
   oHoja:Cells( nRow, 20):Set("HorizontalAlignment",3)
   oHoja:Cells( nRow, 20):Font:Bold := .T.
   nRow++
   oHoja:Cells( nRow, 08):Value :="RESUMEN"
   oHoja:Cells( nRow, 08):Font:Bold := .T.
   oHoja:Cells( nRow, 14):Value :="Base Imponible"
   oHoja:Cells( nRow, 14):Set("HorizontalAlignment",3)
   oHoja:Cells( nRow, 14):Font:Bold := .T.
   oHoja:Cells( nRow, 15):Value :="Débito Fiscal"
   oHoja:Cells( nRow, 15):Set("HorizontalAlignment",3)
   oHoja:Cells( nRow, 15):Font:Bold := .T.
   oHoja:Cells( nRow, 18):Value :="el Comprador"
   oHoja:Cells( nRow, 18):Set("HorizontalAlignment",3)
   oHoja:Cells( nRow, 18):Font:Bold := .T.
   oHoja:Cells( nRow, 20):Value :="Percibido"
   oHoja:Cells( nRow, 20):Set("HorizontalAlignment",3)
   oHoja:Cells( nRow, 20):Font:Bold := .T.
   nRow++
   nRowIni2:=nRowIni
   nRowIni:=nRow
   nRowFin2:=nRowFin
   FOR nI := 1 TO LEN(aResumen)
       oHoja:Cells( nRow, 8):Font:Size := 8
       IF nI=1
          oHoja:Cells( nRow, 8):Value := "Total ventas Internas No Gravadas"
       ENDIF
       IF nI=2
          oHoja:Cells( nRow, 8):Value := "Sum. de las Ventas de Exportación"
       ENDIF
       IF nI=3
          oHoja:Cells( nRow, 8):Value := "Sum. de las Ventas Internas Afectas Solo Alicuota general"
       ENDIF
       IF nI=4
          oHoja:Cells( nRow, 8):Value := "Sum. de las Ventas Internas Afectas en Alicuota general+Adicional"
       ENDIF
       IF nI=5
          oHoja:Cells( nRow, 8):Value := "Sum. de las Ventas Internas Afectas en Alicuota Reducida"
       ENDIF
       IF nI=1 .OR. nI=2
          oHoja:Cells( nRow, 14):Value := aResumen[nI,1]
          oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
          oHoja:Cells( nRow, 15):Value := aResumen[nI,2]
          oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
          oHoja:Cells( nRow, 18):Value := aResumen[nI,3]
          oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
          oHoja:Cells( nRow, 20):Value := aResumen[nI,4]
          oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
       ELSE
          cLetra1:=oHojaSuma:Name+"!"+IIF( nI=3,"B",IIF( nI=4,"C" ,"D" ) )
          cLetra2:=oHojaSuma:Name+"!"+IIF( nI=3,"E",IIF( nI=4,"F" ,"G" ) )
          cLetra3:=oHojaSuma:Name+"!"+IIF( nI=3,"H",IIF( nI=4,"I" ,"J" ) )
          cLetra4:=oHojaSuma:Name+"!"+IIF( nI=3,"K",IIF( nI=4,"L" ,"M" ) )
          cRange2:=cLetra1+ALLTRIM(STR(INT(nRowIni2)))+":"+cLetra1+ALLTRIM(STR(INT(nRowFin)))
          oHoja:Cells( nRow, 14):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
          oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
          cRange2:=cLetra2+ALLTRIM(STR(INT(nRowIni2)))+":"+cLetra2+ALLTRIM(STR(INT(nRowFin)))
          oHoja:Cells( nRow, 15):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
          oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
          cRange2:=cLetra3+ALLTRIM(STR(INT(nRowIni2)))+":"+cLetra3+ALLTRIM(STR(INT(nRowFin)))
          oHoja:Cells( nRow, 18):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
          oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
          cRange2:=cLetra4+ALLTRIM(STR(INT(nRowIni2)))+":"+cLetra4+ALLTRIM(STR(INT(nRowFin)))
          oHoja:Cells( nRow, 20):Value := IIF( lPrimera,0 ,"=+"+cFunSuma+"("+cRange2+")")
          oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
       ENDIF
       nRowFin:=nRow
       nRow++
   NEXT
   cRange1:="H"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Range( cRange1 ):Borders():LineStyle  := 1
   cRange1:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Range( cRange1 ):Borders():LineStyle  := 1
   cRange1:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Range( cRange1 ):Borders():LineStyle  := 1
   cRange1:="N"+ALLTRIM(STR(INT(nRowIni)))+":N"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 14):Value := "=+"+cFunSuma+"("+cRange1+")"
   oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
   IF lSombra2
      oHoja:Cells( nRow, 14):Interior:Color := 12632256
   ENDIF
   cRange1:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 15):Value := "=+"+cFunSuma+"("+cRange1+")"
   oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
   IF lSombra2
      oHoja:Cells( nRow, 15):Interior:Color := 12632256
   ENDIF
   cRange1:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 18):Value := "=+"+cFunSuma+"("+cRange1+")"
   oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
   IF lSombra2
      oHoja:Cells( nRow, 18):Interior:Color := 12632256
   ENDIF
   cRange1:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
   oHoja:Cells( nRow, 20):Value := "=+"+cFunSuma+"("+cRange1+")"
   oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
   IF lSombra2
      oHoja:Cells( nRow, 20):Interior:Color := 12632256
   ENDIF
   nRow+=2
   oHojaX:=oHoja
   IF lSepAju
      oHojaX:=oHoja2
      oHoja2:Cells:Font:Name := "Arial"
      oHoja2:Cells:Font:Size := 8
      oHoja2:PageSetup:Set("Zoom",nZoom)
      oHoja2:PageSetup:Set("Orientation",2 )
      oHoja2:PageSetup:Set("PaperSize" , 1 )
      oHoja2:PageSetup:Set("PrintTitleRows","A1:A"+ALLTRIM(STR(nDinamic)))
      IF lNumera
         oHoja2:PageSetup:Set("RightHeader", "&P" )
      ENDIF
      nRow:=1
      FOR nI := 1 TO LEN(aCab1)
          oHoja2:Cells( nRow, 1):Value := aCab1[nI,1]
          oHoja2:Cells( nRow, 1 ):Font:Size := IIF( nI=1,12 ,10 )
          nRow++
      NEXT
      FOR nI := 1 TO LEN(aRangos)
          oHoja2:Cells( nRow, 1):Value := aRangos[nI,1]
          oHoja2:Cells( nRow, 1 ):Font:Size := 8
          oHoja2:Cells( nRow, 1 ):Font:Bold := .T.
          nRow++
      NEXT
      nRow++
      cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":T"+ALLTRIM(STR(INT(nRow+1)))
      oHoja2:Range( cRange1 ):Borders():LineStyle  := 1
      oHoja2:Cells(nRow,1):Value:="AJUSTES"
      oHoja2:Cells( nRow, 1 ):Font:Size := 10
      oHoja2:Cells( nRow, 1 ):Font:Bold := .T.
      oHoja2:Cells(nRow,16):Value:="VENTAS INTERNAS O"
      oHoja2:Range( "P"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
      oHoja2:Cells( nRow, 16):Set("HorizontalAlignment",3)
      IF lSombra1
         oHoja2:Range( cRange1 ):Interior:Color := 12632256
      ENDIF
      nRow++
      oHoja2:Cells(nRow,16):Value:="EXPORTACIONES GRAVADAS"
      oHoja2:Range( "P"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
      oHoja2:Cells( nRow, 16):Set("HorizontalAlignment",3)
      cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":T"+ALLTRIM(STR(INT(nRow+1)))
      IF lSombra1
         oHoja2:Range( cRange1 ):Interior:Color := 12632256
      ENDIF
      nRow++
      nRowIni:=nRow
      nRowIniCab:=nRow
      nJ:=1
      FOR nI := 1 TO LEN(aCabecera)
          IF nI=16 .OR. nI=32
             oHoja2:Range( "P"+ALLTRIM(STR(nRow))+":R"+ALLTRIM(STR(nRow))):Merge()
             oHoja2:Cells( nRow, 16):Set("HorizontalAlignment",3)
          ENDIF
          oHoja2:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/3)) ):Value := aCabecera[nI]
          oHoja2:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/3)) ):Set("HorizontalAlignment",3)
          IF LEN(aCabecera)/3=nI/nJ
             nRow++
             nJ++
          ENDIF
      NEXT
      cRange1:="A"+ALLTRIM(STR(INT(nRowIni)))+":V"+ALLTRIM(STR(INT(nRowIni+2)))
      oHoja2:Range( cRange1 ):Borders():LineStyle  := 1
      IF lSombra1
         oHoja2:Range( cRange1 ):Interior:Color := 12632256
      ENDIF
      nRowIni:=nRow
      nRowFin:=nRow
      // RETENCIONES NO DECLARADAS EN SU MOMENTO
      nI:=1
      (cAliRet)->(ORDSETFOCUS(4))
      (cAliFac)->(ORDSETFOCUS(1))
      (cAliRet)->(DBSEEK(DTOS(dFecha1),.T.))
      DO WHILE lRetAju .AND. !(cAliRet)->(EOF()) .AND. (cAliRet)->RTI_FECHA<=dFecha2
         IF (cAliRet)->RTI_FECHA<dFecha1
            (cAliRet)->(DBSKIP())
            LOOP
         ENDIF
         IF !(cAliFac)->(DBSEEK((cAliRet)->RTI_TIPDOC+(cAliRet)->RTI_NUMERO)) .OR. (cAliFac)->FAC_ESTADO="A" .OR. LEFT(DTOS((cAliFac)->FAC_FECHA),6)>=LEFT(DTOS((cAliRet)->RTI_FECHA),6)
            (cAliRet)->(DBSKIP())
            LOOP
         ENDIF
         IF (cAliRet)->RTI_ESTADO="A"
            (cAliRet)->(DBSKIP())
            LOOP
         ENDIF
         cRif:=SPACE(15)
         cRsocial:=SPACE(nLenRs)
         lContri:= .F.
         IF (cAliCli)->(DBSEEK((cAliFac)->FAC_CODCLI)) .AND. (cAliCli)->(TYPE("CLI_COLLIB"))="C"
            lContri:=( (cAliCli)->CLI_COLLIB<>"N" )
         ELSE
            lContri:=( (cAliFac)->FAC_CONTRI<>"N" )
         ENDIF
         cLetraBase:=IIF( !lContri,"R" ,"Q" )
         nBaseCol:=IIF( !lContri,18 ,17 )
         IF (cAliCli)->(DBSEEK((cAliFac)->FAC_CODCLI))
            cRif:=PADR((cAliCli)->CLI_RIF,15)
            cRsocial:=OemToAnsi(PADR((cAliCli)->CLI_NOMBRE,nLenRs))
         ENDIF
         IF (cAliFac)->FAC_NUMMEM>0 .AND. LEFT((cAliFac)->FAC_CODCLI,1)="0" .AND. LEN(ALLTRIM((cAliFac)->FAC_CODCLI))=1 .AND. (cAliMem)->(DBSEEK((cAliFac)->FAC_NUMMEM))
            cRsocial:=OemToAnsi(PADR((cAliMem)->MEM_DESCRI,nLenRs))
            cRif:=PADR(MEMOLINE((cAliMem)->MEM_MEMO,,5),15) // PADR((cAliCli)->CLI_RIF,15)
         ENDIF
         cDocNum:=(cAliFac)->FAC_NUMERO
         cDocFis:=PADR((cAliRet)->RTI_DOCFIS,10)
         cAfec:=PADR((cAliRet)->RTI_FACASO,10)
         nTotDoc:=(cAliFac)->FAC_NETO
         nDescue:=((cAliFac)->FAC_DESCUE*(cAliFac)->FAC_BRUTO/100)
         nRecarg:=((cAliFac)->FAC_RECARG*(cAliFac)->FAC_BRUTO/100)
         nBase:=(cAliFac)->FAC_BRUTO-nDescue+nRecarg+(cAliFac)->FAC_OTROS-(cAliFac)->FAC_EXONER
         nExento:=(cAliFac)->FAC_EXONER
         nIva:=(cAliFac)->FAC_IMPTO+(cAliFac)->FAC_IMPICS+(cAliFac)->FAC_ICS_A2+(cAliFac)->FAC_IVARED
         nAlic:=RATA(nIva,nBase)
         nIvaRet:=(cAliRet)->RTI_MONTO
         nFactor:=1
         IF (cAliFac)->FAC_TIPO="CR"
            nFactor:=-1
         ENDIF
         nTotDoc:=nTotDoc*nFactor
         nBase:=nBase*nFactor
         nExento:=nExento*nFactor
         nIva:=nIva*nFactor
         nIvaRet:=nIvaRet*nFactor
         nIvaPer:=nIva-nIvaRet
         cCadena1:=SPACE(54)
         IF (cAliFac)->FAC_TIPO="FC"
            cCadena1:=PADC(cDocNum,10)+" "+PADC(cDocFis,10)+" "+SPACE(10)+" "+SPACE(10)+" "+PADC(IIF( A->FAC_ESTADO="A","03-anu" ,"01-reg" ),10)
         ENDIF
         IF (cAliFac)->FAC_TIPO="DB"
            cCadena1:=SPACE(10)+" "+PADC(cDocFis,10)+" "+PADC(cDocNum,10)+" "+SPACE(10)+" "+PADC(IIF( A->FAC_ESTADO="A","03-anu" ,"01-reg" ),10)
         ENDIF
         IF (cAliFac)->FAC_TIPO="CR"
            cCadena1:=SPACE(10)+" "+PADC(cDocFis,10)+" "+SPACE(10)+" "+PADC(cDocNum,10)+" "+PADC(IIF( A->FAC_ESTADO="A","03-anu" ,"01-reg" ),10)
         ENDIF
         cCadena2:=cAfec+" "+PADL(ALLTRIM(TRANSFORM(nTotDoc,"999,999,999.99")),14)
         cCadena3:=PADL(ALLTRIM(TRANSFORM(nExento,"999,999,999.99")),14)+" "+PADL(ALLTRIM(TRANSFORM(nBase,"999,999,999.99")),14)
         cCadena4:=PADL(ALLTRIM(TRANSFORM(nIvaRet,"999,999,999.99")),14)+" "+PADL(ALLTRIM(TRANSFORM(nIvaPer,"999,999,999.99")),14)
         oHoja2:Cells( nRow, 01):Value := IIF( lDinamic,'='+cFunCelda+'("ROW"'+cSepFun+'A'+ALLTRIM(STR(nRow))+')-'+ALLTRIM(STR(nDinamic)),nI)
         oHoja2:Cells( nRow, 02):Value := (cAliFac)->FAC_FECHA
         oHoja2:Cells( nRow, 03):Value := ALLTRIM(cRif)
         oHoja2:Cells( nRow, 04):Value := ALLTRIM(cRSocial)
         oHoja2:Cells( nRow, 05):Value := " "
         oHoja2:Cells( nRow, 06):Value := " "
         oHoja2:Cells( nRow, 07):Value := " "
         oHoja2:Cells( nRow, 08):Set( "NumberFormat", "@")
         oHoja2:Cells( nRow, 08):Value := ALLTRIM(LEFT(cCadena1,10))
         oHoja2:Cells( nRow, 09):Set( "NumberFormat", "@")
         oHoja2:Cells( nRow, 09):Value := ALLTRIM(SUBS(cCadena1,12,10))
         oHoja2:Cells( nRow, 10):Set( "NumberFormat", "@")
         oHoja2:Cells( nRow, 10):Value := ALLTRIM(SUBS(cCadena1,23,10))
         oHoja2:Cells( nRow, 11):Set( "NumberFormat", "@")
         oHoja2:Cells( nRow, 11):Value := ALLTRIM(SUBS(cCadena1,34,10))
         oHoja2:Cells( nRow, 12):Set( "NumberFormat", "@")
         oHoja2:Cells( nRow, 12):Value := ALLTRIM(RIGHT(cCadena1,10))
         oHoja2:Cells( nRow, 14):Value := nTotDoc
         oHoja2:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
         oHoja2:Cells( nRow, nBaseCol):Value := nBase
         oHoja2:Cells( nRow, nBaseCol):Set( "NumberFormat", cMascMon )
         oHoja2:Cells( nRow, 19):Value := nAlic/100
         oHoja2:Cells( nRow, 19):Set( "NumberFormat", cMascPor )
         oHoja2:Cells( nRow, 20):Value := nIva
         oHoja2:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
         oHoja2:Cells( nRow, 21):Value := nIvaRet
         oHoja2:Cells( nRow, 21):Set( "NumberFormat", cMascMon )
         nRowFin:=nRow
         cRange1:="A"+ALLTRIM(STR(INT(nRow)))+":V"+ALLTRIM(STR(INT(nRow)))
         oHoja2:Range( cRange1 ):Borders():LineStyle  := 1
         nRow++
         nI++
         (cAliRet)->(DBSKIP())
      ENDDO
      //
      nRow++
      cRange1:="M"+ALLTRIM(STR(INT(nRow)))+":V"+ALLTRIM(STR(INT(nRow)))
      oHoja2:Range( cRange1 ):Borders():LineStyle  := 1
      oHoja2:Cells( nRow, 13):Value := "Totales"
      oHoja2:Cells( nRow, 13 ):Font:Bold := .T.
      cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":N"+ALLTRIM(STR(INT(nRowFin)))
      oHoja2:Cells( nRow, 14):Value := IIF( .F. ,0 ,"=+"+cFunSuma+"("+cRange2+")")
      oHoja2:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
      cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
      oHoja2:Cells( nRow, 15):Value := IIF( .F. ,0 ,"=+"+cFunSuma+"("+cRange2+")")
      oHoja2:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
      cRange2:="P"+ALLTRIM(STR(INT(nRowIni)))+":P"+ALLTRIM(STR(INT(nRowFin)))
      oHoja2:Cells( nRow, 16):Value := IIF( .F. ,0 ,"=+"+cFunSuma+"("+cRange2+")")
      oHoja2:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
      cRange2:="Q"+ALLTRIM(STR(INT(nRowIni)))+":Q"+ALLTRIM(STR(INT(nRowFin)))
      oHoja2:Cells( nRow, 17):Value := IIF( .F. ,0 ,"=+"+cFunSuma+"("+cRange2+")")
      oHoja2:Cells( nRow, 17):Set( "NumberFormat", cMascMon )
      cRange2:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
      oHoja2:Cells( nRow, 18):Value := IIF( .F. ,0 ,"=+"+cFunSuma+"("+cRange2+")")
      oHoja2:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
      cRange2:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
      oHoja2:Cells( nRow, 20):Value := IIF( .F. ,0 ,"=+"+cFunSuma+"("+cRange2+")")
      oHoja2:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
      cRange2:="U"+ALLTRIM(STR(INT(nRowIni)))+":U"+ALLTRIM(STR(INT(nRowFin)))
      oHoja2:Cells( nRow, 21):Value := IIF( .F. ,0 ,"=+"+cFunSuma+"("+cRange2+")")
      oHoja2:Cells( nRow, 21):Set( "NumberFormat", cMascMon )
      cRange2:="V"+ALLTRIM(STR(INT(nRowIni)))+":V"+ALLTRIM(STR(INT(nRowFin)))
      oHoja2:Cells( nRow, 22):Value := IIF( .F. ,0 ,"=+"+cFunSuma+"("+cRange2+")")
      oHoja2:Cells( nRow, 22):Set( "NumberFormat", cMascMon )
      cRange2:="M"+ALLTRIM(STR(INT(nRow)))+":V"+ALLTRIM(STR(INT(nRow)))
      IF lSombra2
         oHoja2:Range( cRange2 ):Interior:Color := 12632256
      ENDIF
      nRow+=2
   ENDIF
   nRowIni:=nRow
   oHojaX:Cells( nRow, 13):Font:Size := 8
   oHojaX:Cells( nRow, 13):Value := "Saldo Anterior de Ajuste"
   oHojaX:Cells( nRow, 15):Value := 0.00
   oHojaX:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
   nRow++
   oHojaX:Cells( nRow, 13):Font:Size := 8
   oHojaX:Cells( nRow, 13):Value := "Ajustes de este Período"
   oHojaX:Cells( nRow, 15):Value := 0.00
   oHojaX:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
   nRow++
   oHojaX:Cells( nRow, 13):Font:Size := 8
   oHojaX:Cells( nRow, 13):Value := "Total Ajuste Aplicable"
   oHojaX:Cells( nRow, 16):Value := 0.00
   oHojaX:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
   nRow++
   oHojaX:Cells( nRow, 13):Font:Size := 8
   oHojaX:Cells( nRow, 13):Value := "Ajuste Aplicado"
   oHojaX:Cells( nRow, 16):Value := 0.00
   oHojaX:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
   nRow++
   oHojaX:Cells( nRow, 13):Font:Size := 8
   oHojaX:Cells( nRow, 13):Value := "Saldo de Ajuste próximo período"
   oHojaX:Cells( nRow, 16):Value := "=+P"+ALLTRIM(STR(INT(nRow-2)))+"-P"+ALLTRIM(STR(INT(nRow-1)))
   oHojaX:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
   cRange1:="M"+ALLTRIM(STR(INT(nRowIni)))+":P"+ALLTRIM(STR(INT(nRow)))
   oHojaX:Range( cRange1 ):Borders():LineStyle  := 1
   cRange1:="M"+ALLTRIM(STR(INT(nRow)))+":P"+ALLTRIM(STR(INT(nRow)))
   IF lSombra2
      oHojaX:Range( cRange1 ):Interior:Color := 12632256
   ENDIF
   FOR nI := 1 TO 20
       IF (nI>1 .AND. nI<8) .OR. (nI>13 .AND. nI<19) .OR. nI>19
          oHoja:Columns( nI ):AutoFit()
       ENDIF
   NEXT
   oHoja:Columns( 01 ):Set("ColumnWidth",3.14)
   oHoja:Columns( 08 ):Set("ColumnWidth",6.71)
   oHoja:Columns( 09 ):Set("ColumnWidth",6.71)
   oHoja:Columns( 10 ):Set("ColumnWidth",6.71)
   oHoja:Columns( 11 ):Set("ColumnWidth",6.71)
   oHoja:Columns( 12 ):Set("ColumnWidth",4.57)
   oHoja:Columns( 13 ):Set("ColumnWidth",6.71)
   oHoja:Columns( 19 ):Set("ColumnWidth",5.57)
   IF lSepAju
      FOR nI := 1 TO 20
          IF (nI>1 .AND. nI<8) .OR. (nI>13 .AND. nI<19) .OR. nI>19
             oHoja2:Columns( nI ):AutoFit()
          ENDIF
      NEXT
      oHoja2:Columns( 01 ):Set("ColumnWidth",3.14)
      oHoja2:Columns( 08 ):Set("ColumnWidth",6.71)
      oHoja2:Columns( 09 ):Set("ColumnWidth",6.71)
      oHoja2:Columns( 10 ):Set("ColumnWidth",6.71)
      oHoja2:Columns( 11 ):Set("ColumnWidth",6.71)
      oHoja2:Columns( 12 ):Set("ColumnWidth",4.57)
      oHoja2:Columns( 13 ):Set("ColumnWidth",6.71)
      oHoja2:Columns( 19 ):Set("ColumnWidth",5.57)
      *oHoja2:Cells( 1, 1 ):Select()
   ENDIF
   *oHoja:Cells( 1, 1 ):Select()
   oHoja:End()
   oHoja:=NIL
   oHojaSuma:End()
   oHojaSuma:=NIL
   IF lSepAju
      oHoja2:End()
      oHoja2:=NIL
   ENDIF
   oExcel:End()
   oExcel:=NIL
   AEVAL(aAlias,{|x,y| IIF( !EMPTY(aAlias[y]),(aAlias[y])->(DBCLOSEAREA()),NIL)})
RETURN
