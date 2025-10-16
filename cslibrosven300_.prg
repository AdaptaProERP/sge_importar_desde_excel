// Programa   : CSLIBROSVEN3001
// Fecha/Hora : 05/12/2007 15:41:31
// Propósito  : Emite el Libro de Ventas con salida EXCEL
// Aplicación : Ventas
// Tipo       : STD00000
// Creado Por : Daniel Ramírez
// Observación: Llamado por CSLIBROSVEN30
// Modificaciones : (DR20110201a) Se agrega al Libro la posibilidad de que salgan los Números de
//                                Factura de Contingencia
//                  (DR20110201b) Se ordena por Control Fiscal
// Modificaciones : (DR20110606a) Se amplia el formato "general" paraque incluya al número del comprobante
//                  (DR20110606b) Se eliminan las columnas "Número de Documento" y "Número de Control" cuando se trata de retenciones IVA.
//                  (DR20110606c) Se corrige BUG en asignación de Número de Retención.
#include "dpxbase.ch"

PROCE MAIN(oVen)
   LOCAL oPrn,oFont,oFontT,oFontB,oFontD,aAlias:={}
   LOCAL nCantLines:=0,nCantCols:=0,nColStep:=0,nRowStep:=0,nLinea:=0,nTc:=0,nPage:=0,nRow:=0,nCol:=0,nCounter:=0
   LOCAL cBmp:="",cEstatus,oPen,oShadow,aCabecera,aRangos,cCab,aCab,aCab1,aCab2,lHubo:=.F.,lPrimera:=.T.,aMatriz,nWfont,nHfont,nBaseRed
   LOCAL aPie,aCondicion,nTotal,nMonto,nWid,nHei,nlin2,nI,nJ,nColIni,nColFin,aDet,nX,cImp,nFactor,nPrimCol:=6,nSalida
   LOCAL cRif,cRsocial,cPlanilla,cCadena1,cCadena2,cCadena3,cCadena4,cDocFis,cAfec,nTotDoc,nExento,nBase,nAlic,nIva,nIvaRet,nIvaPer,nDescue,nRecarg,cDocNum,cLetra1,cLetra2,cLetra3,cLetra4
   LOCAL nTotal1,nTotal2,nTotal3,nTotal4,nTotal5,nTotal6
   LOCAL nInternaNG,nExporta,nTotGen,nTotGenAdi,nTotReduc,nIvaGen,nIvaGenAdi,nIvaReduc,nRetGen,nRetGenAdi,nRetReduc,nPerGen,nPerGenAdi,nPerReduc,lRetAju:=.F.
   LOCAL aLim,nCol1,nCol2,nCol3,nCol4,nCol5,nCol6,nCol7,nCol8,nCol9
   LOCAL aLim2,nCol21,nCol22,nCol23
   LOCAL nDemoMax:=2,nDemoLin:=1,cLetraBase,nBaseCol
   LOCAL oExcel,oHoja,oHoja2,oHojax,oHojaSuma,oLanguage,nRowIni,nRowIni2,nRowFin,nRowFin2,cRange1,cRange2,cRange3,cRange4,aResumen:=ARRAY(5,4),nRowInicab,lContri
   // Revisión 03/03/2006
   LOCAL cEmpRif:="",nDinamic
   // Compatibilidad
   LOCAL cMascPor,cMascMon
   // Método Copy&Paste (DR20060326)
   LOCAL cClip,cClipSuma,cSuma27,oClip,nPegaIni,nPegaIniSu,oRange,oTexto
   // Adaptaciones DataPro Windows 
   LOCAL cSql,oTable,cEmpresa,cWhere,nIva,nFecha,nNumero,nCodigo,nTipo,aLine,nAt,aTipIva,aTipDoc
   LOCAL cClave,I,nLine,nRecord,nRecAct,nBaseDet,nMontoRti,cNumRti,nIvaIcs,nIvaGen,nIvaIcs2,cSqlIva
   LOCAL oTableIva,nIvaRed,nIvaRetOri,cEstado,dFechaDoc,dFechaRet,cCliEnOtra,nNeto,cDocTipo,cLibTra
   LOCAL oTableDoc,lDocAnu,nBruto,dFecha1,dFecha2,cRange,lEspanol,nBase15,nBase8,nPor8,nImp8,nImp15
   LOCAL nBaseCol8,cLetraBas8,nUltCol,dFchRti,cFacRti,dFchFacRti,cTipFacRti,cSucFacRti,lRti,nNetRti
   LOCAL cFisFacRti,nNetFacRti,cNumDocRti
   LOCAL aAbc:={"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD"}
   LOCAL cDocTipRti // DR20110606c
   LOCAL nNumFis // DR20110606d

   cClip:=""
   cClipSuma:=""
   cSuma27:=""
   nUltCol:=1
   DO CASE
      CASE oVen:nVForm=1
           nUltCol:=23
      CASE oVen:nVForm=2
           nUltCol:=16
      CASE oVen:nVForm=3
           nUltCol:=21
      CASE oVen:nVForm=4
           nUltCol:=27
   ENDCASE
   IF oVen:lVColRti
      nUltCol++
   ENDIF
   //

   cEmpresa:=oDp:cEmpresa
   dFecha1:=CTOD("01/"+STRZERO(oVen:nMes,2)+"/"+STRZERO(oVen:nAno,4))
   dFecha2:=FCHFINMES(dFecha1)

   cWhere:=""
   IF !(oVen:oModelo:nAt==LEN(oVen:oModelo:aItems))
      cWhere:=" (DOC_MODFIS"+GetWhere("=",oVen:cModelo)+" OR DOC_TIPDOC='RTI')"
   ENDIF

   FOR nI := 1 TO LEN(aResumen)
       aResumen[nI,1]:=""
       aResumen[nI,2]:=""
       aResumen[nI,3]:=""
       aResumen[nI,4]:=""
   NEXT

   lRetAju:=IIF(!oVen:lVSepAju,.F.,oVen:lVRetAju )

   IF oVen:nVForm=1 .OR. oVen:nVForm=4
      nDinamic:=IIF( oVen:lVRifCab,10 ,9 )
   ENDIF
   IF oVen:nVForm=2
      nDinamic:=IIF( oVen:lVRifCab,7 ,6 )
   ENDIF
   IF oVen:nVForm=3
      nDinamic:=IIF( oVen:lVRifCab,8 ,7 )
   ENDIF
   

   aTipDoc:=ASQL("SELECT TDC_TIPO FROM DPTIPDOCCLI WHERE TDC_LIBVTA=1")
   aTipIva:=ASQL("SELECT TIP_CODIGO FROM DPIVATIP")

   // Hacemos un SELECT sin SUM para poder contar el detalle

   // DR20110201a. Se agrega condición para FCT
   // DR20110201b. Se ordena por Control Fiscal
   cSql:="SELECT DOC_NETO  ,"+;
         "       DOC_FECHA ,"+;
         "       DOC_DESTIN,"+;
         "       RTI_DOCNUM,"+;
         "       RTI_PORCEN,"+;
         "       DOC_CODIGO,"+;
         "       DOC_CODSUC,"+;
         "       IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_TIPPER,CLI_TIPPER)            AS CLI_TIPPER ,"+;
         "       IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_RIF   ,DPCLIENTES.CLI_RIF   ) AS CLI_RIF    ,"+;
         "       IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_NOMBRE,DPCLIENTES.CLI_NOMBRE) AS CLI_NOMBRE ,"+;
         "       IF(DOC_TIPDOC='FAV' OR DOC_TIPDOC='FCT',DOC_NUMERO              ,SPACE(10)             ) AS DOC_FACTURA,"+;
         "       IF(DOC_TIPDOC='DEB'       ,DOC_NUMERO              ,SPACE(10)             ) AS DOC_DEBITO ,"+;
         "       IF(DOC_TIPDOC='DEV'       ,DOC_NUMERO              ,SPACE(10)             ) AS DOC_CREDITO,"+;
         "       IF(DOC_TIPDOC='TIK'       ,DOC_NUMERO              ,DOC_NUMFIS            ) AS DOC_NUMFIS ,"+;
         "       DOC_TIPDOC, "+;
         "       DOC_NUMERO, "+;
         "       DOC_FACAFE, "+;
         "       MOV_IVA   , "+;
         "       MOV_TIPIVA, "+;
         "       DOC_DCTO  , "+;
         "       DOC_RECARG, "+;
         "       DOC_OTROS , "+;
         "       DOC_ESTADO, "+;
         "       DOC_ANUFIS, "+;
         "       IF(DOC_ANUFIS=0,TDC_LIBTRA,'00-REG') AS TDC_LIBTRA, "+;
         "       TDC_LIBVTA, "+;
         "       DOC_PLAEXP, "+;
         "       DOC_CXC   , "+;
         "       IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_CONTRI,CLI_CONTRI)            AS CLI_CONTRI ,"+;
         "       RTI_FECHA , "+;
         "       RTI_DOCTIP, "+;
         "       RTI_NUMERO, "+;
         "       RTI_CODSUC, "+;
         "       CLI_ENOTRA, "+;
         "       DOC_ACT,    "+;
         "       MOV_TOTAL, MOV_IMPOTR"+;
         " FROM DPDOCCLI "+;
         " INNER JOIN DPCLIENTES  ON DOC_CODIGO=CLI_CODIGO "+;
         " INNER JOIN DPTIPDOCCLI ON DOC_TIPDOC=TDC_TIPO   "+;
         " LEFT  JOIN DPCLIENTESCERO ON DOC_CODSUC=CCG_CODSUC AND "+;
         "            DOC_TIPDOC=CCG_TIPDOC AND "+;
         "            DOC_NUMERO=CCG_NUMDOC "+;
         " LEFT  JOIN DPMOVINV ON MOV_CODSUC=DOC_CODSUC AND MOV_TIPDOC=DOC_TIPDOC AND MOV_CODCTA=DOC_CODIGO"+;
         "            AND MOV_DOCUME=DOC_NUMERO AND MOV_INVACT=1 "+;
         " LEFT JOIN DPIVATIP    ON MOV_TIPIVA=TIP_CODIGO   "+;
         " LEFT JOIN DPDOCCLIRTI ON DOC_CODSUC=RTI_CODSUC AND "+;
         "                          DOC_TIPDOC=RTI_TIPDOC AND "+;
         "                          DOC_NUMERO=RTI_NUMERO AND "+;
         "                          DOC_TIPTRA=RTI_TIPTRA "+;
         " WHERE "+GetWhereAnd("DOC_FECHA",dFecha1,dFecha2)+" AND "+;
         "   DOC_CODSUC "+GetWhere("=",oVen:cCodSuc)+;
         "   AND DOC_TIPTRA='D' "+;
         "   AND " + GetWhereOr("DOC_TIPDOC",aTipDoc)+" OR DOC_TIPDOC='TIK' OR DOC_TIPDOC='RTI'"+;
         "   AND " + GetWhereAnd("DOC_FECHA",dFecha1,dFecha2)+;
         "   AND (TDC_LIBVTA=1 OR DOC_TIPDOC='RTI') " +;
         IIF(!empty(cWhere),"       AND "+ cWhere,"")+;
         " ORDER BY DOC_FECHA,DOC_NUMFIS,DOC_NUMERO,MOV_IVA "

   oTable:=OpenTable(cSql,.T.)

   cEmpRif:=oDp:cRif

   // Comentarios de comandos validos
/*
   msgstop(oExcel:WorkSheets(1))
   msgstop(oExcel:WorkSheets(1):name)
   oHoja2:=oExcel:Sheets:Add()
   oHoja := oExcel:Get( "ActiveSheet" )
   msgstop(oExcel:Name)
 */
   oExcel := TOleAuto():New( "Excel.Application" )
   IF oVen:lEDetIdio
      oLanguage:=oExcel:Get("LanguageSettings")
      lEspanol:=EJECUTAR("CSISEXCELSPA",oLanguage:LanguageID(1))
   ENDIF
   IF oVen:lEDetSep
      oVen:cESepDec:=oExcel:Get("International",3)
      oVen:cESepMil:=oExcel:Get("International",4)
      oVen:cESepLis:=oExcel:Get("International",5)
   ENDIF
   cMascPor:="##0"+oVen:cESepDec+"00%"
   cMascMon:="###"+oVen:cESepMil+"###"+oVen:cESepMil+"##0"+oVen:cESepDec+"00"
   IF !oVen:lVSepMil
      cMascMon:="########0"+oVen:cESepDec+"00"
   ENDIF

   IF oVen:lEDetIdio
      oVen:cEFunCel:="CELDA"
      oVen:cEFunSum:="SUMA"
      oVen:cEFunSi:="SI"
      IF !lEspanol
         oVen:cEFunCel:="CELL"
         oVen:cEFunSum:="SUM"
         oVen:cEFunSi:="IF"
      ENDIF
   ENDIF

   oExcel:WorkBooks:Add()
   oHojaSuma:=oExcel:WorkSheets:Add()
   oHojaSuma:Name:="Valores"
   IF oVen:lVSepAju
      oHoja2:=oExcel:WorkSheets:Add()
      oHoja2:Name:="Ajustes"
   ENDIF
   oHoja:=oExcel:WorkSheets:Add()
   oHoja:Name:="Ventas"
//   oExcel:Visible := .T.
   oHoja:Cells:Font:Name := "Arial"
   oHoja:Cells:Font:Size := 8
   oHoja:PageSetup:Set("Zoom",oVen:nVPorAju)
   oHoja:PageSetup:Set("Orientation",2 )
   oHoja:PageSetup:Set("PaperSize" , 1 )
   oHoja:PageSetup:Set("PrintTitleRows","A1:A"+ALLTRIM(STR(nDinamic)))
   IF oVen:lVNumPag
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
   AADD(aCab1,{ALLTRIM(cEmpresa),oFontT})
   IF oVen:lVRifCab
      AADD(aCab1,{ALLTRIM(cEmpRif),oFontT})
   ENDIF
   AADD(aCab1,{"Libro de Ventas",oFontT})

   aCabecera:={}
   IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
      //Primera Linea
      AADD(aCabecera,space(7))
      AADD(aCabecera,padc('Fecha',11))
      AADD(aCabecera,space(15))
      AADD(aCabecera,space(oVen:nVLenRs))
      AADD(aCabecera,space(1))
      AADD(aCabecera,Padc('Núm.Planilla',15))
      AADD(aCabecera,space(1))
      AADD(aCabecera,padc("Número",10))
      AADD(aCabecera,space(10))
      AADD(aCabecera,padc("Número",10))
      AADD(aCabecera,space(10))
      AADD(aCabecera,space(10))
      AADD(aCabecera,padc("Número de",10))
      AADD(aCabecera,padc("Total Ventas",15))
      AADD(aCabecera,padc("Ventas",15))
      IF oVen:nVForm=1 .OR. oVen:nVForm=4
         AADD(aCabecera,"")
         IF oVen:nVForm=1
            AADD(aCabecera,"")
         ENDIF
         IF oVen:nVForm=4
            AADD(aCabecera,"** Alicuota General **")
         ENDIF
         AADD(aCabecera,"")
      ENDIF
      IF oVen:nVForm=3
         AADD(aCabecera,space(15))
      ENDIF
      AADD(aCabecera,space(7))
      AADD(aCabecera,space(15))
      IF oVen:nVForm=4
         AADD(aCabecera,"** Alicuota Reducida **")
         AADD(aCabecera,"")
         AADD(aCabecera,space(7))
         AADD(aCabecera,space(15))
      ENDIF
      AADD(aCabecera,padc("IVA Retenido",15))
      AADD(aCabecera,space(15))
      IF oVen:lVColRti
         AADD(aCabecera,"Número de")
      ENDIF
      //Segunda Linea
      AADD(aCabecera,padc("Oper",7))
      AADD(aCabecera,padc('de la',11))
      AADD(aCabecera,space(15))
      AADD(aCabecera,space(oVen:nVLenRs))
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
      IF oVen:nVForm=1 .OR. oVen:nVForm=4
         IF oVen:nVForm=4
            AADD(aCabecera,"")
         ENDIF
         AADD(aCabecera,"Base Imponible")
         AADD(aCabecera,"")
         IF oVen:nVForm=1
            AADD(aCabecera,"")
         ENDIF
      ENDIF
      IF oVen:nVForm=3
         AADD(aCabecera,space(15))
      ENDIF
      AADD(aCabecera,padc("%",7))
      AADD(aCabecera,padc("Impuesto",15))
      IF oVen:nVForm=4
         AADD(aCabecera,"Base Imponible")
         AADD(aCabecera,"")
         AADD(aCabecera,padc("%",7))
         AADD(aCabecera,padc("Impuesto",15))
      ENDIF
      AADD(aCabecera,padc("Por el",15))
      AADD(aCabecera,padc("IVA",15))
      IF oVen:lVColRti
         AADD(aCabecera,"Comprobante")
      ENDIF
      //Tercera Linea
      AADD(aCabecera,padc("Nro.",7))
      AADD(aCabecera,padc('Factura',11))
      AADD(aCabecera,padc("R.I.F.",15))
      AADD(aCabecera,padc('Nombre o Razón Social',oVen:nVLenRs))
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
      IF oVen:nVForm=1 .OR. oVen:nVForm=4
         AADD(aCabecera,"Exportaciones")
         AADD(aCabecera,"Contribuyentes")
         AADD(aCabecera,"No Contribuyentes")
      ENDIF
      IF oVen:nVForm=3
         AADD(aCabecera,padc("Base imponible",15))
      ENDIF
      AADD(aCabecera,padc("Alicuota",7))
      AADD(aCabecera,padc("IVA",15))
      IF oVen:nVForm=4
         AADD(aCabecera,"Contribuyentes")
         AADD(aCabecera,"No Contribuyentes")
         AADD(aCabecera,padc("Alicuota",7))
         AADD(aCabecera,padc("IVA",15))
      ENDIF
      AADD(aCabecera,padc("Comprador",15))
      AADD(aCabecera,padc("Percibido",15))
      IF oVen:lVColRti
         AADD(aCabecera,"Retención")
      ENDIF
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
      IF oVen:nVForm=1 .OR. oVen:nVForm=4
         IF oVen:nVForm=1
            cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":T"+ALLTRIM(STR(INT(nRow+1)))
         ENDIF
         IF oVen:nVForm=4
            cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":X"+ALLTRIM(STR(INT(nRow+1)))
         ENDIF
         oHoja:Range( cRange1 ):Borders():LineStyle  := 1
         oHoja:Cells(nRow,16):Value:="VENTAS INTERNAS O"
         IF oVen:nVForm=1
            oHoja:Range( "P"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
         ENDIF
         IF oVen:nVForm=4
            oHoja:Range( "P"+ALLTRIM(STR(nRow))+":X"+ALLTRIM(STR(nRow))):Merge()
         ENDIF
         oHoja:Cells( nRow, 16):Set("HorizontalAlignment",3)
         IF oVen:lVSombCab
            oHoja:Range( cRange1 ):Interior:Color := 12632256
         ENDIF
         nRow++
         oHoja:Cells(nRow,16):Value:="EXPORTACIONES GRAVADAS"
         IF oVen:nVForm=1
            oHoja:Range( "P"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
         ENDIF
         IF oVen:nVForm=4
            oHoja:Range( "P"+ALLTRIM(STR(nRow))+":X"+ALLTRIM(STR(nRow))):Merge()
         ENDIF
         oHoja:Cells( nRow, 16):Set("HorizontalAlignment",3)
         IF oVen:nVForm=1
            cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":T"+ALLTRIM(STR(INT(nRow+1)))
         ENDIF
         IF oVen:nVForm=4
            cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":X"+ALLTRIM(STR(INT(nRow+1)))
         ENDIF
         IF oVen:lVSombCab
            oHoja:Range( cRange1 ):Interior:Color := 12632256
         ENDIF
         nRow++
      ENDIF
      nRowIni:=nRow
      nRowIniCab:=nRow
      nJ:=1
      FOR nI := 1 TO LEN(aCabecera)
          IF oVen:nVForm=1
             IF nI=16 .OR. nI=32
                oHoja:Range( "P"+ALLTRIM(STR(nRow))+":R"+ALLTRIM(STR(nRow))):Merge()
                oHoja:Cells( nRow, 16):Set("HorizontalAlignment",3)
             ENDIF
          ENDIF
          IF oVen:nVForm=4
             IF nI=17
                oHoja:Range( "Q"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
                oHoja:Cells( nRow, 17):Set("HorizontalAlignment",3)
             ENDIF
             IF nI=43
                oHoja:Range( "Q"+ALLTRIM(STR(nRow))+":R"+ALLTRIM(STR(nRow))):Merge()
                oHoja:Cells( nRow, 17):Set("HorizontalAlignment",3)
             ENDIF
             IF nI=21
                oHoja:Range( "U"+ALLTRIM(STR(nRow))+":X"+ALLTRIM(STR(nRow))):Merge()
                oHoja:Cells( nRow, 21):Set("HorizontalAlignment",3)
             ENDIF
             IF nI=47
                oHoja:Range( "U"+ALLTRIM(STR(nRow))+":V"+ALLTRIM(STR(nRow))):Merge()
                oHoja:Cells( nRow, 21):Set("HorizontalAlignment",3)
             ENDIF
          ENDIF
          oHoja:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/3)) ):Value := aCabecera[nI]
          oHoja:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/3)) ):Set("HorizontalAlignment",3)
          IF LEN(aCabecera)/3=nI/nJ
             nRow++
             nJ++
          ENDIF
      NEXT
      cRange1:="A"+ALLTRIM(STR(INT(nRowIni)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRowIni+2)))
      oHoja:Range( cRange1 ):Borders():LineStyle  := 1
      IF oVen:lVSombCab
         oHoja:Range( cRange1 ):Interior:Color := 12632256
      ENDIF
   ENDIF
   IF oVen:nVForm=2
      //Primera Linea
      AADD(aCabecera,"FECHA")
      AADD(aCabecera,"NUMERO")
      AADD(aCabecera,"CLIENTE")
      AADD(aCabecera,"NUMERO DE RIF")
      AADD(aCabecera,"TOTAL VENTAS")
      AADD(aCabecera,"VENTAS")
      AADD(aCabecera,"")
      AADD(aCabecera,"VENTAS EXPORTACION")
      AADD(aCabecera,"")
      AADD(aCabecera,"")
      AADD(aCabecera,"VENTAS A NO CONTRIBUYENTES")
      AADD(aCabecera,"")
      AADD(aCabecera,"")
      AADD(aCabecera,"VENTAS A CONTRIBUYENTES")
      AADD(aCabecera,"")
      IF oVen:lVColRti
         AADD(aCabecera,"COMPROBANTE DE")
      ENDIF
      //Segunda Linea
      AADD(aCabecera,"")
      AADD(aCabecera,"FACTURA")
      AADD(aCabecera,"")
      AADD(aCabecera,"")
      AADD(aCabecera,"INCLUYE I.V.A.")
      AADD(acabecera,"EXENTAS")
      AADD(acabecera,"BASE")
      AADD(acabecera,"%")
      AADD(aCabecera,"IVA")
      AADD(aCabecera,"BASE")
      AADD(aCabecera,"%")
      AADD(aCabecera,"IVA")
      AADD(aCabecera,"BASE")
      AADD(aCabecera,"%")
      AADD(aCabecera,"IVA")
      IF oVen:lVColRti
         AADD(aCabecera,"RETENCION")
      ENDIF
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
      nRowIni:=nRow
      nRowIniCab:=nRow
      nJ:=1
      FOR nI := 1 TO LEN(aCabecera)
          IF nI=9
             oHoja:Range( "G"+ALLTRIM(STR(nRow))+":I"+ALLTRIM(STR(nRow))):Merge()
             oHoja:Cells( nRow, 09):Set("HorizontalAlignment",3)
          ENDIF
          IF nI=12
             oHoja:Range( "J"+ALLTRIM(STR(nRow))+":L"+ALLTRIM(STR(nRow))):Merge()
             oHoja:Cells( nRow, 12):Set("HorizontalAlignment",3)
          ENDIF
          IF nI=15
             oHoja:Range( "M"+ALLTRIM(STR(nRow))+":O"+ALLTRIM(STR(nRow))):Merge()
             oHoja:Cells( nRow, 15):Set("HorizontalAlignment",3)
          ENDIF
          oHoja:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/2)) ):Value := aCabecera[nI]
          oHoja:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/2)) ):Set("HorizontalAlignment",3)
          IF LEN(aCabecera)/2=nI/nJ
             nRow++
             nJ++
          ENDIF
      NEXT
      cRange1:="A"+ALLTRIM(STR(INT(nRowIni)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRowIni+1)))
      oHoja:Range( cRange1 ):Borders():LineStyle  := 1
      IF oVen:lVSombCab
         oHoja:Range( cRange1 ):Interior:Color := 12632256
      ENDIF
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
   nPegaIni:=nRow // Copy&Paste
   nPegaIniSu:=nRow // Copy&Paste

   oTable:Replace("DOC_OPERAC",STRZERO(0,6)) // Número de Transacción
   oTable:Replace("DOC_FACAFE",SPACE(10))    // Factura Afectada
   oTable:Replace("DOC_MTOIVA",0        )    // Alicuota IVA
   oTable:Replace("DOC_EXONER",0        )    // Exento
   oTable:Replace("DOC_MTORTI",0        )    // Monto de Retención RTI

   oTable:Gotop()

   nIva   := oTable:FieldPos("MOV_IVA")
   nFecha := oTable:FieldPos("DOC_FECHA")
   nNumero:= oTable:FieldPos("DOC_NUMERO")
   nCodigo:= oTable:FieldPos("DOC_CODIGO")
   nTipo  := oTable:FieldPos("DOC_TIPDOC")

   // Busca los Impuestos de los Documentos
   WHILE !oTable:Eof() 
      IF Empty(oTable:MOV_TIPIVA) .AND. oTable:DOC_TIPDOC<>'RTI' // No tiene Productos
         aLine:=ACLONE(oTable:aDataFill[oTable:Recno()])
         EJECUTAR("DPDOCCLIIVA",oTable:DOC_CODSUC,oTable:DOC_TIPDOC,oTable:DOC_CODIGO,oTable:DOC_NUMERO,.F.,;
                               oTable:DOC_DCTO  ,oTable:DOC_RECARG,oTable:DOC_OTROS,NIL,"V")

         WHILE .T.
           nAt:=ASCAN(oDp:aArrayIva,{|a,n|a[4]=0})
           IF nAt=0
             EXIT
           ENDIF
           oDp:aArrayIva:=ARREDUCE(oDp:aArrayIva,nAt)
         ENDDO

         FOR I:=1 TO LEN(oDp:aArrayIva)
           nAt:=ASCAN(aTipIva,{|a,n|a[1]==oDp:aArrayIva[I,1]})

           IF nAt>0 .AND. oDp:aArrayIva[I,4]<>0 .AND. LEN(oDp:aArrayIva)>0
             aLine[oTable:FieldPos("MOV_TIPIVA")]:=oDp:aArrayIva[I,1]
             aLine[oTable:FieldPos("MOV_TOTAL" )]:=oDp:aArrayIva[I,4]
             aLine[oTable:FieldPos("MOV_IVA"   )]:=oDp:aArrayIva[I,3]

             IF I=1
               oTable:aDataFill[oTable:Recno()]:=ACLONE(aLine)
             ELSE
               AADD(oTable:aDataFill,ACLONE(aLine))
             ENDIF
           ENDIF
         NEXT I
      ENDIF
      oTable:DbSkip()
   ENDDO
   oTable:GoTop()
   oVen:oMeter:SetTotal(oTable:RecCount())
//oTable:Browse()

   DO WHILE !oTable:Eof()

      cDocNum:=oTable:DOC_NUMERO
      nFactor:=oTable:DOC_CXC
      lPrimera:=.F.
      cRif:=SPACE(15)
      cRsocial:=SPACE(oVen:nVLenRs)
      cPlanilla:=oTable:DOC_PLAEXP
      cDocFis:=PADR(oTable:DOC_NUMFIS,10)
      cAfec:=PADR(oTable:DOC_FACAFE,10)
      nTotDoc:=oTable:DOC_NETO*oTable:DOC_ACT
      lRti:=(oTable:DOC_TIPDOC="RTI")

      oVen:oMeter:Set(oTable:RecNo())

      oVen:oSayRecord:SetText("Registro:"+LSTR(oTable:Recno())+"/"+LSTR(oTable:Reccount()))

      IF !oVen:lVLibRti .AND. lRti
         oTable:Skip()
         LOOP
      ENDIF

      cFacRti:=""
      cTipFacRti:=""
      cSucFacRti:=""
      cFisFacRti:=""
      nNetFacRti:=0
      dFchFacRti:=CTOD("")
      cNumDocRti:=""
      cDocTipRti:="" // DR20110606c
      IF lRti
         cDocTipRti:=oTable:DOC_TIPDOC // DR20110606c
         SQLGET("DPDOCCLIRTI","RTI_NUMERO,RTI_TIPDOC,RTI_CODSUC","RTI_CODSUC"+GetWhere("=",oTable:DOC_CODSUC)+" AND RTI_DOCTIP"+GetWhere("=",oTable:DOC_TIPDOC)+" AND RTI_DOCNUM"+GetWhere("=",oTable:DOC_NUMERO)+" AND RTI_TIPTRA='D'")
         //  MsgInfo("RTI_CODSUC"+GetWhere("=",oTable:DOC_CODSUC)+" AND RTI_DOCTIP"+GetWhere("=",oTable:DOC_TIPDOC)+" AND RTI_DOCNUM"+GetWhere("=",oTable:DOC_NUMERO)+" AND RTI_TIPTRA='D'")
         IF LEN(oDp:aRow)>0
            cFacRti:=oDp:aRow[1]
            cTipFacRti:=oDp:aRow[2]
            cSucFacRti:=oDp:aRow[3]
         ENDIF
         IF EMPTY(cFacRti)
            oTable:Skip()
            LOOP
         ENDIF
         SQLGET("DPDOCCLI","DOC_FECHA,DOC_FACAFE,DOC_NETO*DOC_ACT,DOC_ESTADO,DOC_ANUFIS","DOC_CODSUC"+GetWhere("=",cSucFacRti)+" AND DOC_TIPDOC"+GetWhere("=",cTipFacRti)+" AND DOC_NUMERO"+GetWhere("=",cFacRti)+" AND DOC_TIPTRA='D'")
         IF LEN(oDp:aRow)>0
            dFchFacRti:=oDp:aRow[1]
            cFisFacRti:=oDp:aRow[2]
            nNetFacRti:=oDp:aRow[3]
            IF oDp:aRow[4]="NU" .OR. oDp:aRow[5]
               oTable:Skip()
               LOOP
            ENDIF
         ENDIF
         // MsgInfo("DOC_CODSUC"+GetWhere("=",cSucFacRti)+" AND DOC_TIPDOC"+GetWhere("=",cTipFacRti)+" AND DOC_NUMERO"+GetWhere("=",cFacRti)+" AND DOC_TIPTRA='D'")
         // Es del mismo período y no va como RTI, sino desde el documento retenido
         IF EMPTY(dFchFacRti) .OR. LEFT(DTOS(dFchFacRti),6)=LEFT(DTOS(oTable:DOC_FECHA),6)
            oTable:Skip()
            LOOP
         ENDIF
         cDocNum:=cFacRti
         cAfec:=cFisFacRti
         nTotDoc:=0
         cNumDocRti:=oTable:DOC_NUMERO
      ENDIF
      
      cClave:=oTable:DOC_TIPDOC+oTable:DOC_NUMERO+oTable:DOC_CODIGO

      nLine  :=0
      nRecord:=oTable:Recno()       // Primer Renglón

      nBase:=0
      nExento:=0
      nIva:=0
      nAlic:=0
      nIvaRet:=0

      cPlanilla:=oTable:DOC_PLAEXP
      cAfec:=oTable:DOC_FACAFE
      cNumRti:=IIF(EMPTY(oTable:RTI_DOCNUM).OR.lRti,"",oTable:RTI_DOCNUM)
      nIvaRetOri:=0
      dFchRti:=CTOD("")
      IF !EMPTY(cNumRti)
         SQLGET("DPDOCCLI","DOC_NETO*DOC_ACT,DOC_FECHA","DOC_CODSUC"+GetWhere("=",oTable:RTI_CODSUC)+" AND DOC_TIPDOC"+GetWhere("=",oTable:RTI_DOCTIP)+" AND DOC_NUMERO"+GetWhere("=",oTable:RTI_DOCNUM)+" AND DOC_TIPTRA='D'")
         IF LEN(oDp:aRow)>0
            nIvaRetOri:=oDp:aRow[1]
            dFchRti:=oDp:aRow[2]
         ENDIF
         // Si está encendida 'Retenciones en Libro' entonces debe ser del mismo período la retención
         IF oVen:lVLibRti .AND. LEFT(DTOS(dFchRti),6)<>LEFT(DTOS(oTable:DOC_FECHA),6)
            cNumRti:=""
            nIvaRetOri:=0
            dFchRti:=CTOD("")
         ENDIF
      ENDIF
      nIvaIcs:=0
      nIvaGen:=0
      nIvaIcs2:=0
      nIvaRed:=0
      lDocAnu:=oTable:DOC_ANUFIS
      cEstado:=oTable:DOC_ESTADO
      dFechaDoc:=oTable:DOC_FECHA
      dFechaRet:=oTable:RTI_FECHA
      cRif:=PADR(oTable:CLI_RIF,15)
      cRsocial:=PADR(oTable:CLI_NOMBRE,oVen:nVLenRs)
      cCliEnOtra:=oTable:CLI_ENOTRA
      nNeto:=IIF(lRti,0,oTable:DOC_NETO*oTable:DOC_ACT)
      cDocTipo:=IIF(lRti,cTipFacRti,oTable:DOC_TIPDOC)
      cLibTra:=IIF(lRti,SQLGET("DPTIPDOCCLI","TDC_LIBTRA","TDC_TIPO"+GetWhere("=",cDocTipo)),oTable:TDC_LIBTRA)
      nExento:=0 // IIF(lRti,0,oTable:DOC_EXONER*oTable:DOC_ACT)
      nNetRti:=IIF(lRti,oTable:DOC_NETO*oTable:DOC_ACT,0)*-1
      lContri:=oTable:CLI_CONTRI<>"N"
      nImp15:=0
      nImp8:=0
      nPor8:=0
      nBase8:=0
      nBase15:=0

      WHILE !oTable:Eof() .AND. cClave=oTable:DOC_TIPDOC+oTable:DOC_NUMERO+oTable:DOC_CODIGO
         IF !lRti
            IF oTable:DOC_ANUFIS
               oTable:Replace("MOV_TOTAL",0)
            ENDIF
            oDp:nDesc     :=oTable:DOC_DCTO*oTable:DOC_ACT
            oDp:nRecarg   :=oTable:DOC_RECARG*oTable:DOC_ACT
            oDp:nDocOtros :=oTable:DOC_OTROS*oTable:DOC_ACT
            oDp:nBruto    :=oTable:MOV_TOTAL*oTable:DOC_ACT
            oDp:nMtoDesc  :=Porcen(oDp:nBruto,oDp:nDesc  )*oTable:DOC_ACT
            oDp:nMtoRecarg:=Porcen(oDp:nBruto,oDp:nRecarg)*oTable:DOC_ACT
            oDp:nNeto     :=oDp:nBruto+oDp:nMtoRecarg-oDp:nMtoDesc+oDp:nDocOtros
            oDp:nMtoVar   :=oDp:nBruto-oDp:nNeto
            oDp:nPorVar   :=(100-RATA(oDp:nNeto,oDp:nBruto))*-1
            oTable:Replace("MOV_BASE"  ,(oTable:MOV_TOTAL+PORCEN(oTable:MOV_TOTAL,oDp:nPorVar))*oTable:DOC_ACT)
            oTable:Replace("DOC_MTOIVA",PORCEN(oTable:MOV_BASE,oTable:MOV_IVA)*oTable:DOC_ACT)

            nLine:=nLine+1
            IF nLine=1 .AND. !Empty(cNumRti)
               nMontoRti:=SQLGET("DPDOCCLI","DOC_NETO*DOC_ACT","DOC_CODSUC"+GetWhere("=",oVen:cCodSuc)+" AND "+;
                                                       "DOC_TIPDOC"+GetWhere("=","RTI")            +" AND "+;
                                                       "DOC_NUMERO"+GetWhere("=",oTable:RTI_DOCNUM)+" AND "+;
                                                       "DOC_CODIGO"+GetWhere("=",oTable:DOC_CODIGO))

               oTable:Replace("DOC_MTORTI",nMontoRti)
               nIvaRet:=nIvaRet+nMontoRti
            ENDIF
            IF nLine=1 .AND. oTable:MOV_IVA=0
               nExento:=nExento+(oTable:DOC_NETO*oTable:DOC_ACT)
               oTable:Replace("DOC_EXONER",oTable:DOC_NETO*oTable:DOC_ACT)
               oTable:Replace("MOV_BASE"  ,0) // No va para el Libro de Venta
            ENDIF

            IF nLine>1
              oTable:Replace("DOC_NETO",0)
              nBaseDet :=oTable:MOV_BASE*oTable:DOC_ACT        // EXENTO
 
              IF oTable:MOV_IVA=0            // EXENTO
                 oTable:Replace("MOV_BASE",0) // No va para el Libro de Venta
                 oTable:Replace("DOC_NETO",0) //
                 nRecAct:=oTable:Recno()
                 oTable:Goto(nRecord)         // Registro Anterior
//                 nExento:=nExento+(oTable:DOC_NETO*oTable:DOC_ACT)
                 oTable:Replace("DOC_EXONER",nBaseDet*oTable:DOC_ACT)
                 oTable:Goto(nRecAct)
              ENDIF
            ENDIF
            IF oTable:DOC_ANUFIS
               oTable:Replace("MOV_BASE"  ,0)
               oTable:Replace("MOV_IVA"   ,0)
               oTable:Replace("DOC_EXONER",0)   
               oTable:Replace("DOC_MTOIVA",0)
               oTable:Replace("DOC_MTORTI",0)
            ENDIF
            nBase:=nBase+(oTable:MOV_BASE*oTable:DOC_ACT)
            nIva:=nIva+PORCEN(oTable:MOV_BASE,oTable:MOV_IVA)*oTable:DOC_ACT
            IF !EMPTY(oTable:MOV_TIPIVA)
               cSqlIva:="SELECT * FROM DPIVATABC WHERE CTI_FECHA"+GetWhere("<=",oTable:DOC_FECHA)+" AND CTI_TIPO"+GetWhere("=",oTable:MOV_TIPIVA)+" ORDER BY CTI_FECHA DESC LIMIT 1"
               oTableIva:=OpenTable(cSqlIva)
               IF oTableIva:RecCount()>0 .AND. oTable:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="GN"
                  nIvaGen:=nIvaGen+PORCEN(oTable:MOV_BASE,oTable:MOV_IVA)*oTable:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTable:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="S1"
                  nIvaIcs:=nIvaIcs+PORCEN(oTable:MOV_BASE,oTable:MOV_IVA)*oTable:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTable:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="S2"
                  nIvaIcs2:=nIvaIcs2+PORCEN(oTable:MOV_BASE,oTable:MOV_IVA)*oTable:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTable:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="RD"
                  nIvaRed:=nIvaRed+PORCEN(oTable:MOV_BASE,oTable:MOV_IVA)*oTable:DOC_ACT
                  nBase8:=nBase8+oTable:MOV_BASE*oTable:DOC_ACT
                  nPor8:=oTableIva:CTI_VENTA
               ENDIF
            ENDIF
         ELSE
            // Buscamos el IVA
            cSql:="SELECT "+;
                  "DOC_CODSUC, "+;
                  "DOC_TIPDOC, "+;
                  "DOC_NUMERO, "+;
                  "DOC_FECHA,  "+;
                  "DOC_ANUFIS, "+;
                  "DOC_NUMFIS, "+;
                  "CLI_CODIGO, "+;
                  "MOV_IVA   , "+;
                  "MOV_TIPIVA, "+;
                  "DOC_CXC   , "+;
                  "DOC_ACT   , "+;
                  "MOV_TOTAL , "+;
                  "MOV_IMPOTR, "+;
                  "DOC_DCTO  , "+;
                  "DOC_RECARG, "+;
                  "DOC_OTROS,  "+;
                  "DOC_NETO,   "+;
                  "IF(DOC_ANUFIS=0,TDC_LIBTRA,'03-Anu') AS TDC_LIBTRA, "+;
                  "IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_RIF   ,DPCLIENTES.CLI_RIF   ) AS CLI_RIF    ,"+;
                  "IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_NOMBRE,DPCLIENTES.CLI_NOMBRE) AS CLI_NOMBRE ,"+;
                  "IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_CONTRI,DPCLIENTES.CLI_CONTRI) AS CLI_CONTRI  "+;
                  " FROM DPDOCCLI INNER JOIN DPCLIENTES ON CLI_CODIGO=DOC_CODIGO "+;
                  " INNER JOIN DPTIPDOCCLI ON DOC_TIPDOC=TDC_TIPO   "+;
                  " LEFT  JOIN DPCLIENTESCERO ON DOC_CODSUC=CCG_CODSUC AND "+;
                  "            DOC_TIPDOC=CCG_TIPDOC AND "+;
                  "            DOC_NUMERO=CCG_NUMDOC "+;
                  " LEFT  JOIN DPMOVINV ON MOV_CODSUC=DOC_CODSUC AND MOV_TIPDOC=DOC_TIPDOC AND MOV_CODCTA=DOC_CODIGO"+;
                  "            AND MOV_DOCUME=DOC_NUMERO AND MOV_INVACT=1 "+;
                  "WHERE DOC_CODSUC"+GetWhere("=",cSucFacRti)+" AND DOC_TIPDOC"+GetWhere("=",cTipFacRti)+;
                  " AND DOC_NUMERO"+GetWhere("=",cDocNum)+" AND DOC_TIPTRA='D'"

            oTableDoc:=OpenTable(cSql)
            oTableDoc:GoTop()
            oTableDoc:Replace("DOC_OPERAC",STRZERO(0,6)) // Número de Transacción
            oTableDoc:Replace("DOC_FACAFE",SPACE(10))    // Factura Afectada
            oTableDoc:Replace("DOC_MTOIVA",0        )    // Alicuota IVA
            oTableDoc:Replace("DOC_EXONER",0        )    // Exento
            oTableDoc:Replace("DOC_MTORTI",0        )    // Monto de Retención RTI

            lContri:= oTableDoc:CLI_CONTRI<>"N"
            cLetraBase:=IIF( !lContri,"R" ,"Q" )
            nBaseCol:=IIF( !lContri,18 ,17 )
            cLetraBas8:=IIF( !lContri,"V" ,"U" )
            nBaseCol8:=IIF( !lContri,22 ,21 )
            cRif:=PADR(oTableDoc:CLI_RIF,15)
            cRsocial:=PADR(oTableDoc:CLI_NOMBRE,oVen:nVLenRs)
            cDocNum:=oTableDoc:DOC_NUMERO
            cDocFis:=PADR(oTableDoc:DOC_NUMFIS,10)
            cAfec:=PADR(oTableDoc:DOC_FACAFE,10)
            nTotDoc:=oTableDoc:DOC_NETO*oTableDoc:DOC_CXC*oTableDoc:DOC_ACT
            nBruto:=0
            IF Empty(oTableDoc:MOV_TIPIVA) // No tiene Productos
               aLine:=ACLONE(oTableDoc:aDataFill[oTableDoc:Recno()])
               EJECUTAR("DPDOCCLIIVA",oTableDoc:DOC_CODSUC,oTableDoc:DOC_TIPDOC,oTableDoc:CLI_CODIGO,oTableDoc:DOC_NUMERO,.F.,;
                                  oTableDoc:DOC_DCTO  ,oTableDoc:DOC_RECARG,oTableDoc:DOC_OTROS,NIL,"V")
               WHILE .T.
                 nAt:=ASCAN(oDp:aArrayIva,{|a,n|a[4]=0})
                 IF nAt=0
                   EXIT
                 ENDIF
                 oDp:aArrayIva:=ARREDUCE(oDp:aArrayIva,nAt)
               ENDDO
               FOR I:=1 TO LEN(oDp:aArrayIva)
                   nAt:=ASCAN(aTipIva,{|a,n|a[1]==oDp:aArrayIva[I,1]})
                   IF nAt>0 .AND. oDp:aArrayIva[I,4]<>0 .AND. LEN(oDp:aArrayIva)>0
                      aLine[oTableDoc:FieldPos("MOV_TIPIVA")]:=oDp:aArrayIva[I,1]
                      aLine[oTableDoc:FieldPos("MOV_TOTAL" )]:=oDp:aArrayIva[I,4]
                      aLine[oTableDoc:FieldPos("MOV_IVA"   )]:=oDp:aArrayIva[I,3]
                      IF I=1
                         oTableDoc:aDataFill[oTableDoc:Recno()]:=ACLONE(aLine)
                      ELSE
                         AADD(oTableDoc:aDataFill,ACLONE(aLine))
                      ENDIF
                   ENDIF
               NEXT I
            ENDIF
            nLine:=0
            nBase:=0
            nIva:=0
            nIvaGen:=0
            nIvaIcs:=0
            nIvaIcs2:=0
            nIvaRed:=0
            nImp15:=0
            nImp8:=0
            nPor8:=0
            nBase8:=0
            nBase15:=0
            oTableDoc:GoTop()
            WHILE !oTableDoc:Eof()
               IF oTableDoc:DOC_ANUFIS
                  oTableDoc:Replace("MOV_TOTAL",0)
               ENDIF
               oDp:nDesc     :=oTableDoc:DOC_DCTO*oTableDoc:DOC_ACT
               oDp:nRecarg   :=oTableDoc:DOC_RECARG*oTableDoc:DOC_ACT
               oDp:nDocOtros :=oTableDoc:DOC_OTROS*oTableDoc:DOC_ACT
               oDp:nBruto    :=oTableDoc:MOV_TOTAL*oTableDoc:DOC_ACT
               oDp:nMtoDesc  :=Porcen(oDp:nBruto,oDp:nDesc  )*oTableDoc:DOC_ACT
               oDp:nMtoRecarg:=Porcen(oDp:nBruto,oDp:nRecarg)*oTableDoc:DOC_ACT
               oDp:nNeto     :=oDp:nBruto+oDp:nMtoRecarg-oDp:nMtoDesc+oDp:nDocOtros
               oDp:nMtoVar   :=oDp:nBruto-oDp:nNeto
               oDp:nPorVar   :=(100-RATA(oDp:nNeto,oDp:nBruto))*-1
               oTableDoc:Replace("MOV_BASE"  ,(oTableDoc:MOV_TOTAL+PORCEN(oTableDoc:MOV_TOTAL,oDp:nPorVar))*oTableDoc:DOC_ACT)
               oTableDoc:Replace("DOC_MTOIVA",PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT)
               nLine:=nLine+1
               IF nLine=1 .AND. oTableDoc:MOV_IVA=0
                  oTableDoc:Replace("DOC_EXONER",oTableDoc:DOC_NETO*oTableDoc:DOC_ACT)
                  oTableDoc:Replace("MOV_BASE"  ,0) // No va para el Libro de Venta
               ENDIF
               IF nLine>1
                  oTableDoc:Replace("DOC_NETO",0)
                  nBaseDet :=oTableDoc:MOV_BASE*oTableDoc:DOC_ACT        // EXENTO
                  IF oTableDoc:MOV_IVA=0            // EXENTO
                     oTableDoc:Replace("MOV_BASE",0) // No va para el Libro de Venta
                     oTableDoc:Replace("DOC_NETO",0) //
                     nRecAct:=oTableDoc:Recno()
                     oTableDoc:Goto(nRecord)         // Registro Anterior
                     oTableDoc:Replace("DOC_EXONER",nBaseDet*oTableDoc:DOC_ACT)
                     oTableDoc:Goto(nRecAct)
                  ENDIF
               ENDIF
               nBase:=nBase+(oTableDoc:MOV_BASE*oTableDoc:DOC_ACT)
               nIva:=nIva+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               cSqlIva:="SELECT * FROM DPIVATABC WHERE CTI_FECHA"+GetWhere("<=",oTableDoc:DOC_FECHA)+" AND CTI_TIPO"+GetWhere("=",oTableDoc:MOV_TIPIVA)+" ORDER BY CTI_FECHA DESC LIMIT 1"
               oTableIva:=OpenTable(cSqlIva)
               IF oTableIva:RecCount()>0 .AND. oTableDoc:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="GN"
                  nIvaGen:=nIvaGen+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTableDoc:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="S1"
                  nIvaIcs:=nIvaIcs+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTableDoc:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="S2"
                  nIvaIcs2:=nIvaIcs2+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTableDoc:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="RD"
                  nIvaRed:=nIvaRed+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
                  nBase8:=nBase8+oTableDoc:MOV_BASE*oTableDoc:DOC_ACT
                  nPor8:=PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               ENDIF
               oTableDoc:Skip()
            ENDDO
            nImp15:=nIva+nIvaIcs+nIvaIcs2
            nImp8:=nIvaRed
            nBase15:=nBase-nBase8
            oTableDoc:GoTop()
            nExento:=oTableDoc:DOC_EXONER
            nAlic:=RATA(nIva,nBase)
            nFactor:=oTableDoc:DOC_CXC
            nTotDoc:=nTotDoc*nFactor
            nBase:=nBase*nFactor
            nExento:=nExento*nFactor
            nIva:=nIva*nFactor
            nBase8:=nBase8*nFactor
            nBase15:=nBase15*nFactor
            nImp8:=nImp8*nFactor
            nImp15:=nImp15*nFactor
            nIvaPer:=nIva-nIvaRet
            oTableDoc:End()
         ENDIF
         oTable:DbSkip()
      ENDDO

      nAlic:=RATA(nIva,nBase)
      IF oVen:nVForm=1
         cLetraBase:=IIF( !lContri,"R" ,"Q" )
         nBaseCol:=IIF( !lContri,18 ,17 )
      ENDIF
      IF oVen:nVForm=2
         cLetraBase:=IIF( !lContri,"J" ,"M" )
         nBaseCol:=IIF( !lContri,10 ,13 )
      ENDIF
      IF oVen:nVForm=4
         cLetraBase:=IIF( !lContri,"R" ,"Q" )
         nBaseCol:=IIF( !lContri,18 ,17 )
         cLetraBas8:=IIF( !lContri,"V" ,"U" )
         nBaseCol8:=IIF( !lContri,22 ,21 )
      ENDIF
      nImp15:=nIva+nIvaIcs+nIvaIcs2
      nImp8:=nIvaRed
      nBase15:=nBase-nBase8
      IF !EMPTY(cNumRti) .AND. (!lRetAju .OR. LEFT(DTOS(dFechaDoc),6)>=LEFT(DTOS(dFechaRet),6))
         nIvaRet:=nIvaRetOri
         IF !lDocAnu
            nRetGen+=IIF( nIva<>0 .AND. nIvaIcs=0 .AND. nIvaIcs2=0,(nIvaRetOri*nFactor) , 0)
            nRetGenAdi+=IIF( nIva<>0 .AND. (nIvaIcs<>0 .OR. nIvaIcs2<>0),(nIvaRetOri*nFactor) , 0)
            nRetReduc+=IIF( nIvaRed<>0,(nIvaRetOri*nFactor) , 0)
         ENDIF
      ENDIF

      IF lRti
         nIvaRet:=nNetRti*-1
         cNumRti:=SQLGET("DPDOCCLIRTI","RTI_NUMCLI","RTI_CODSUC"+GEtWhere("=",cSucFacRti)+" AND RTI_DOCTIP"+GetWhere("=",cDocTipRti)+" AND RTI_DOCNUM"+GetWhere("=",cNumDocRti)+" AND RTI_TIPTRA='D'") // DR20110606c. cNumDocRti
      ENDIF

      IF cCliEnOtra="E" .AND. nIva=0
         nExento:=0
         IF oVen:nVForm=1 .OR. oVen:nVForm=4
            cLetraBase:="P"
            nBaseCol:=16
         ENDIF
         IF oVen:nVForm=2
            cDocNum:=IIF( cDocTipo="FC",PADR("N/A",08),cDocNum)
            cDocFis:=IIF( cDocTipo="FC",PADR("N/A",10),cDocNum)
            cLetraBase:="G"
            nBaseCol:=7
         ENDIF
         nExporta+=(nNeto*nFactor)
      ENDIF

      nTotDoc:=nTotDoc*nFactor
      nBase:=nBase*nFactor
      nExento:=nExento*nFactor
      nIva:=nIva*nFactor
      nIvaRet:=nIvaRet*nFactor
      nBase8:=nBase8*nFactor
      nBase15:=nBase15*nFactor
      nImp8:=nImp8*nFactor
      nImp15:=nImp15*nFactor

      nIvaPer:=nIva-nIvaRet

      // Copy&Paste
      cSuma27:=""
      IF oVen:nVForm<>2
         DO CASE
            CASE !lDocAnu .AND. nIva<>0 .AND. nIvaIcs=0 .AND. nIvaIcs2=0
                 cClipSuma:=cClipSuma+"VIAG"+CHR(9)
            CASE !lDocAnu .AND. nIva<>0 .AND. (nIvaIcs<>0 .OR. nIvaIcs2<>0)
                 cClipSuma:=cClipSuma+"VIGA"+CHR(9)
            CASE !lDocAnu .AND. nIvaRed<>0 .AND. nIva<=0
                 IF oVen:nVForm=1 .OR. oVen:nVForm=3
                    cClipSuma:=cClipSuma+"VIAR"+CHR(9)
                 ENDIF
                 IF oVen:nVForm=4
                    cClipSuma:=cClipSuma+" "+CHR(9)
                    cSuma27:="VIAR"
                 ENDIF
            OTHERWISE
                 cClipSuma:=cClipSuma+" "+CHR(9)
         ENDCASE
      ENDIF
      IF lRti
         nTotDoc:=0
         nBase:=0
         nExento:=0
         nIva:=0
         nBase8:=0
         nBase15:=0
         nImp8:=0
         nImp15:=0
         nIvaPer:=nIvaRet*-1
      ENDIF
      //
      IF oVen:nVForm=1 .OR. oVen:nVForm=3
         IF !lDocAnu .AND. nIva<>0 .AND. nIvaIcs=0 .AND. nIvaIcs2=0
            aResumen[3,1]:=IIF( EMPTY(aResumen[3,1]),"=" ,aResumen[3,1])
            aResumen[3,2]:=IIF( EMPTY(aResumen[3,2]),"=" ,aResumen[3,2])
            aResumen[3,3]:=IIF( EMPTY(aResumen[3,3]),"=" ,aResumen[3,3])
            aResumen[3,4]:=IIF( EMPTY(aResumen[3,4]),"=" ,aResumen[3,4])
            IF oVen:nVForm=1
               aResumen[3,1]:=aResumen[3,1]+"+"+cLetraBase+ALLTRIM(STR(INT(nRow)))
               aResumen[3,2]:=aResumen[3,2]+"+T"+ALLTRIM(STR(INT(nRow)))
               aResumen[3,3]:=aResumen[3,3]+"+U"+ALLTRIM(STR(INT(nRow)))
               aResumen[3,4]:=aResumen[3,4]+"+V"+ALLTRIM(STR(INT(nRow)))
            ENDIF
            IF oVen:nVForm=3
               aResumen[3,1]:=aResumen[3,1]+"+P"+ALLTRIM(STR(INT(nRow)))
               aResumen[3,2]:=aResumen[3,2]+"+R"+ALLTRIM(STR(INT(nRow)))
               aResumen[3,3]:=aResumen[3,3]+"+S"+ALLTRIM(STR(INT(nRow)))
               aResumen[3,4]:=aResumen[3,4]+"+T"+ALLTRIM(STR(INT(nRow)))
            ENDIF
            *oHojaSuma:Cells( nRow, 01):Value := "VIAG"
         ENDIF

         IF !lDocAnu .AND. nIva<>0 .AND. (nIvaIcs<>0 .OR. nIvaIcs2<>0)
            aResumen[4,1]:=IIF( EMPTY(aResumen[4,1]),"=" ,aResumen[4,1])
            aResumen[4,2]:=IIF( EMPTY(aResumen[4,2]),"=" ,aResumen[4,2])
            aResumen[4,3]:=IIF( EMPTY(aResumen[4,3]),"=" ,aResumen[4,3])
            aResumen[4,4]:=IIF( EMPTY(aResumen[4,4]),"=" ,aResumen[4,4])
            IF oVen:nVForm=1
               aResumen[4,1]:=aResumen[4,1]+"+"+cLetraBase+ALLTRIM(STR(INT(nRow)))
               aResumen[4,2]:=aResumen[4,2]+"+T"+ALLTRIM(STR(INT(nRow)))
               aResumen[4,3]:=aResumen[4,3]+"+U"+ALLTRIM(STR(INT(nRow)))
               aResumen[4,4]:=aResumen[4,4]+"+V"+ALLTRIM(STR(INT(nRow)))
            ENDIF
            IF oVen:nVForm=3
               aResumen[4,1]:=aResumen[4,1]+"+P"+ALLTRIM(STR(INT(nRow)))
               aResumen[4,2]:=aResumen[4,2]+"+R"+ALLTRIM(STR(INT(nRow)))
               aResumen[4,3]:=aResumen[4,3]+"+S"+ALLTRIM(STR(INT(nRow)))
               aResumen[4,4]:=aResumen[4,4]+"+T"+ALLTRIM(STR(INT(nRow)))
            ENDIF
            *oHojaSuma:Cells( nRow, 01):Value := "VIGA"
         ENDIF

         IF !lDocAnu .AND. nIvaRed<>0 .AND. nIva<=0
            aResumen[5,1]:=IIF( EMPTY(aResumen[5,1]),"=" ,aResumen[5,1])
            aResumen[5,2]:=IIF( EMPTY(aResumen[5,2]),"=" ,aResumen[5,2])
            aResumen[5,3]:=IIF( EMPTY(aResumen[5,3]),"=" ,aResumen[5,3])
            aResumen[5,4]:=IIF( EMPTY(aResumen[5,4]),"=" ,aResumen[5,4])
            IF oVen:nVForm=1
               aResumen[5,1]:=aResumen[5,1]+"+"+cLetraBase+ALLTRIM(STR(INT(nRow)))
               aResumen[5,2]:=aResumen[5,2]+"+T"+ALLTRIM(STR(INT(nRow)))
               aResumen[5,3]:=aResumen[5,3]+"+U"+ALLTRIM(STR(INT(nRow)))
               aResumen[5,4]:=aResumen[5,4]+"+V"+ALLTRIM(STR(INT(nRow)))
            ENDIF
            IF oVen:nVForm=3
               aResumen[5,1]:=aResumen[5,1]+"+P"+ALLTRIM(STR(INT(nRow)))
               aResumen[5,2]:=aResumen[5,2]+"+R"+ALLTRIM(STR(INT(nRow)))
               aResumen[5,3]:=aResumen[5,3]+"+S"+ALLTRIM(STR(INT(nRow)))
               aResumen[5,4]:=aResumen[5,4]+"+T"+ALLTRIM(STR(INT(nRow)))
            ENDIF
            *oHojaSuma:Cells( nRow, 01):Value := "VIAR"
         ENDIF
      ENDIF

      IF lDocAnu
         // DR20110201a. Se agrega condición para FCT
         IF cDocTipo="FAV" .OR. cDocTipo="FCT"
            cAfec:=PADR(cDocNum,10)
         ENDIF
         nTotDoc:=0
         nBase:=0
         nExento:=0
         nIva:=0
         nAlic:=0
         nIvaRet:=0
         nIvaPer:=0
         nImp8:=0
         nImp15:=0
         nBase8:=0
         nBase15:=0
      ENDIF

      cCadena1:=SPACE(54)
      // DR20110201a. Se agrega condición para FCT
      IF !lRti // DR20110606b. Se condiciona
         IF cDocTipo="FAV" .OR. cDocTipo="FCT"
            cCadena1:=PADC(cDocNum,10)+" "+PADC(cDocFis,10)+" "+SPACE(10)+" "+SPACE(10)+" "+PADC(cLibTra ,10)
         ENDIF
         IF cDocTipo="DEB"
            cCadena1:=SPACE(10)+" "+PADC(cDocFis,10)+" "+PADC(cDocNum,10)+" "+SPACE(10)+" "+PADC(cLibTra ,10)
         ENDIF
         IF cDocTipo="CRE" .OR.cDocTipo="DEV"
            cCadena1:=SPACE(10)+" "+PADC(cDocFis,10)+" "+SPACE(10)+" "+PADC(cDocNum,10)+" "+PADC(cLibTra ,10)
         ENDIF
      ELSE
         // DR20110606b. Si es RTI, solo el Tipo de Transacción.
         cCadena1:=SPACE(10)+" "+SPACE(10)+" "+SPACE(10)+" "+SPACE(10)+" "+PADC(cLibTra ,10)
         cAfec:=cDocNum
      ENDIF
      cCadena2:=cAfec // +" "+PADL(ALLTRIM(TRANSFORM(nTotDoc,"999,999,999.99")),14)
      cCadena3:=PADL(ALLTRIM(TRANSFORM(nExento,"999,999,999.99")),14)+" "+PADL(ALLTRIM(TRANSFORM(nBase,"999,999,999.99")),14)
      cCadena4:=PADL(ALLTRIM(TRANSFORM(nIvaRet,"999,999,999.99")),14)+" "+PADL(ALLTRIM(TRANSFORM(nIvaPer,"999,999,999.99")),14)
      aDet:={}
      // Copy&Paste
      // Vamos llenando la variable de pegado
      IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
         cClip:=cClip+IIF( oVen:lVConDin,'='+oVen:cEFunCel+'("ROW"'+oVen:cESepLis+'A'+ALLTRIM(STR(nRow))+')-'+ALLTRIM(STR(nDinamic)),STR(nI))+CHR(9)
         cClip:=cClip+DTOC(dFechaDoc)+CHR(9)
         cClip:=cClip+ALLTRIM(cRif)+CHR(9)
         cClip:=cClip+ALLTRIM(cRSocial)+CHR(9)
         cClip:=cClip+" "+CHR(9)
         cClip:=cClip+ALLTRIM(cPlanilla)+CHR(9)
         cClip:=cClip+" "+CHR(9)
         cClip:=cClip+ALLTRIM(LEFT(cCadena1,10))+CHR(9)
         cClip:=cClip+ALLTRIM(SUBS(cCadena1,12,10))+CHR(9)
         cClip:=cClip+ALLTRIM(SUBS(cCadena1,23,10))+CHR(9)
         cClip:=cClip+ALLTRIM(SUBS(cCadena1,34,10))+CHR(9)
         cClip:=cClip+ALLTRIM(RIGHT(cCadena1,10))+CHR(9)
         cClip:=cClip+ALLTRIM(LEFT(cCadena2,10))+CHR(9)
         cClip:=cClip+EJECUTAR("CSSTREXCEL",nTotDoc,oVen:cESepDec,16,2)+CHR(9)
         cClip:=cClip+EJECUTAR("CSSTREXCEL",nExento,oVen:cESepDec,16,2)+CHR(9)
         IF oVen:nVForm=1 .OR. oVen:nVForm=4
            IF nBaseCol=16
               IF oVen:nVForm=1
                  cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase,oVen:cESepDec,16,2)+CHR(9)
               ENDIF
               IF oVen:nVForm=4
                  cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase15,oVen:cESepDec,16,2)+CHR(9)
               ENDIF
               cClip:=cClip+" "+CHR(9)
               cClip:=cClip+" "+CHR(9)
            ENDIF
            IF nBaseCol=17
               cClip:=cClip+" "+CHR(9)
               IF oVen:nVForm=1
                  cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase,oVen:cESepDec,16,2)+CHR(9)
               ENDIF
               IF oVen:nVForm=4
                  cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase15,oVen:cESepDec,16,2)+CHR(9)
               ENDIF
               cClip:=cClip+" "+CHR(9)
            ENDIF
            IF nBaseCol=18
               cClip:=cClip+" "+CHR(9)
               cClip:=cClip+" "+CHR(9)
               IF oVen:nVForm=1
                  cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase,oVen:cESepDec,16,2)+CHR(9)
               ENDIF
               IF oVen:nVForm=4
                  cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase15,oVen:cESepDec,16,2)+CHR(9)
               ENDIF
            ENDIF
         ENDIF
         IF oVen:nVForm=3
            cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase,oVen:cESepDec,16,2)+CHR(9)
         ENDIF
         IF oVen:nVForm=1 .OR. oVen:nVForm=3
            cClip:=cClip+EJECUTAR("CSSTREXCEL",nAlic/100,oVen:cESepDec,16,2)+CHR(9)
            cClip:=cClip+EJECUTAR("CSSTREXCEL",nIva,oVen:cESepDec,16,2)+CHR(9)
         ENDIF
         IF oVen:nVForm=4
            cClip:=cClip+EJECUTAR("CSSTREXCEL",Rata(nImp15,nBase15)/100,oVen:cESepDec,16,2)+CHR(9)
            cClip:=cClip+EJECUTAR("CSSTREXCEL",nImp15,oVen:cESepDec,16,2)+CHR(9)
            IF nBaseCol8=21
               cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase8,oVen:cESepDec,16,2)+CHR(9)
               cClip:=cClip+" "+CHR(9)
            ENDIF
            IF nBaseCol8=22
               cClip:=cClip+" "+CHR(9)
               cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase8,oVen:cESepDec,16,2)+CHR(9)
            ENDIF
            cClip:=cClip+EJECUTAR("CSSTREXCEL",Rata(nImp8,nBase8),oVen:cESepDec,16,2)+CHR(9)
            cClip:=cClip+EJECUTAR("CSSTREXCEL",nImp8,oVen:cESepDec,16,2)+CHR(9)
         ENDIF
         cClip:=cClip+EJECUTAR("CSSTREXCEL",nIvaRet,oVen:cESepDec,16,2)+CHR(9)
         cClip:=cClip+EJECUTAR("CSSTREXCEL",nIvaPer,oVen:cESepDec,16,2)+CHR(9)+IIF(oVen:lVColRti,"",CHR(13))
      ENDIF
      IF oVen:nVForm=2
         cClip:=cClip+DTOC(dFechaDoc)+CHR(9)
         cClip:=cClip+cDocTipo+cDocNum+CHR(9)
         cClip:=cClip+ALLTRIM(cRSocial)+CHR(9)
         cClip:=cClip+ALLTRIM(cRif)+CHR(9)
         cClip:=cClip+EJECUTAR("CSSTREXCEL",nTotDoc,oVen:cESepDec,16,2)+CHR(9)
         cClip:=cClip+EJECUTAR("CSSTREXCEL",nExento,oVen:cESepDec,16,2)+CHR(9)+IIF( nBase<>0,"" ,IIF(oVen:lVColRti,"",CHR(13)) )
         IF nBase<>0
            DO CASE
               CASE nBaseCol=7
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase,oVen:cESepDec,16,2)+CHR(9)
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nAlic/100,oVen:cESepDec,16,2)+CHR(9)
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nIva,oVen:cESepDec,16,2)+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)+IIF(oVen:lVColRti,"",CHR(13))
               CASE nBaseCol=10
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase,oVen:cESepDec,16,2)+CHR(9)
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nAlic/100,oVen:cESepDec,16,2)+CHR(9)
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nIva,oVen:cESepDec,16,2)+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)+IIF(oVen:lVColRti,"",CHR(13))
               CASE nBaseCol=13
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+" "+CHR(9)
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nBase,oVen:cESepDec,16,2)+CHR(9)
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nAlic/100,oVen:cESepDec,16,2)+CHR(9)
                    cClip:=cClip+EJECUTAR("CSSTREXCEL",nIva,oVen:cESepDec,16,2)+CHR(9)+IIF(oVen:lVColRti,"",CHR(13))
            ENDCASE
         ENDIF
      ENDIF
      cClip:=cClip+IIF(oVen:lVColRti,cNumRti+CHR(13),"")
      //
      nRowFin:=nRow
      IF oVen:nVForm=1 .OR. oVen:nVForm=2 .OR. oVen:nVForm=3
         cRange1:="A"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
      ENDIF
      *oHoja:Range( cRange1 ):Borders():LineStyle  := 1
      // VIAG, VIGA o VIRE
      // nBase
      // nIva
      // nIvaRet
      // nIvaPer
      // Copy&Paste
      IF oVen:nVForm=1 
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!'+cLetraBase+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!'+cLetraBase+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!'+cLetraBase+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!U'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!U'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!U'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!V'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!V'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!V'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)+CHR(13)
      //
      ENDIF
      IF oVen:nVForm=3
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!P'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!P'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!P'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!R'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!R'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!R'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!S'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!S'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!S'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)+CHR(13)
      ENDIF
      IF oVen:nVForm=4
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!'+cLetraBase+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!'+cLetraBase+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(AA'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!'+cLetraBas8+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!T'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(AA'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!X'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!Y'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!Y'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(AA'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!Y'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIAG"'+oVen:cESepLis+oHoja:Name+'!Z'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(A'+ALLTRIM(STR(nRow))+'="VIGA"'+oVen:cESepLis+oHoja:Name+'!Z'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+'=+'+oVen:cEFunSi+'(AA'+ALLTRIM(STR(nRow))+'="VIAR"'+oVen:cESepLis+oHoja:Name+'!Z'+ALLTRIM(STR(nRow))+oVen:cESepLis+'0)'+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+" "+CHR(9)
         cClipSuma:=cClipSuma+cSuma27+CHR(9)+CHR(13)
      ENDIF
      nColIni:=nColStep*nPrimCol
      // Copy&Paste
      // Cada 20k llenamos el portapapeles y lo pegamos en la Hoja de Excel
      IF LEN(cClip)>20000
         oClip := TClipBoard():New()
         oClip:Clear()
         oClip:SetText( cClip )
         cRange:="A"+ALLTRIM(STR(nPegaIni))
         IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
            // DR20110606a. Ampliamos el formato "General"
            cRange2:="E"+ALLTRIM(STR(INT(nPegaIni)))+":M"+ALLTRIM(STR(INT(nRow)))
            oHoja:Range(cRange2):Set( "NumberFormat", "@")
         ENDIF
         cRange2:=aAbc[nUltCol]+ALLTRIM(STR(INT(nPegaIni)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
         oHoja:Range(cRange2):Set( "NumberFormat", "@")
         oHoja:Activate()
         oRange:=oHoja:Range( cRange )
         oRange:Select()
         oHoja:Paste()
         oClip:End()
         cClip := ""
         nPegaIni := nRow + 1
      ENDIF
      IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
         IF LEN(cClipSuma)>15000
            oClip := TClipBoard():New()
            oClip:Clear()
            oClip:SetText( cClipSuma )
            cRange:="A"+ALLTRIM(STR(nPegaIniSu))
            oHojasuma:Activate()
            oRange:=oHojaSuma:Range( cRange )
            oRange:Select()
            oHojaSuma:Paste()
            oClip:End()
            cClipSuma := ""
            nPegaIniSu:= nRow + 1
         ENDIF
      ENDIF
      nTotal1+=nTotDoc
      nTotal2+=nExento
      nTotal3+=nBase
      nTotal4+=nIva
      nTotal5+=nIvaRet
      nTotal6+=nIvaPer
      nI++
      nDemoLin++
      nRow++
   ENDDO
   oTable:End()
   // Copy&Paste
   IF LEN(cClip)>0
      oClip := TClipBoard():New()
      oClip:Clear()
      oClip:SetText( cClip )
      cRange:="A"+ALLTRIM(STR(nPegaIni))
      IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
         // DR20110606a. Se amplía el formato.
         cRange2:="E"+ALLTRIM(STR(INT(nPegaIni)))+":M"+ALLTRIM(STR(INT(nRow)))
         oHoja:Range(cRange2):Set( "NumberFormat", "@")
      ENDIF
      cRange2:=aAbc[nUltCol]+ALLTRIM(STR(INT(nPegaIni)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
      oHoja:Range(cRange2):Set( "NumberFormat", "@")
      oHoja:Activate()
      oRange:=oHoja:Range( cRange )
      oRange:Select()
      oHoja:Paste()
      oClip:End()
      cClip := ""
      nPegaIni := nRow + 1
   ENDIF
   IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
      IF LEN(cClipSuma)>0
         oClip := TClipBoard():New()
         oClip:Clear()
         oClip:SetText( cClipSuma )
         cRange:="A"+ALLTRIM(STR(nPegaIniSu))
         oHojasuma:Activate()
         oRange:=oHojaSuma:Range( cRange )
         oRange:Select()
         oHojaSuma:Paste()
         oClip:End()
         cClipSuma := ""
         nPegaIniSu:= nRow + 1
      ENDIF
   ENDIF
   IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
      cRange2:="H"+ALLTRIM(STR(INT(nRowIni)))+":M"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", "@")
      IF oVen:nVForm=1 .OR. oVen:nVForm=4
         cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
      ENDIF
      IF oVen:nVForm=3
         cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":P"+ALLTRIM(STR(INT(nRowFin)))
      ENDIF
      oHoja:Range(cRange2):Set( "NumberFormat", cMascMon)
      IF oVen:nVForm=1 .OR. oVen:nVForm=4
         cRange2:="S"+ALLTRIM(STR(INT(nRowIni)))+":S"+ALLTRIM(STR(INT(nRowFin)))
      ENDIF
      IF oVen:nVForm=3
         cRange2:="Q"+ALLTRIM(STR(INT(nRowIni)))+":Q"+ALLTRIM(STR(INT(nRowFin)))
      ENDIF
      oHoja:Range(cRange2):Set( "NumberFormat", cMascPor)
      IF oVen:nVForm=1 .OR. oVen:nVForm=4
         cRange2:="T"+ALLTRIM(STR(INT(nRowIni)))+":V"+ALLTRIM(STR(INT(nRowFin)))
      ENDIF
      IF oVen:nVForm=3
         cRange2:="R"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
      ENDIF
      oHoja:Range(cRange2):Set( "NumberFormat", cMascMon)
      IF oVen:nVForm=4
         cRange2:="W"+ALLTRIM(STR(INT(nRowIni)))+":W"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Range(cRange2):Set( "NumberFormat", cMascPor)
         cRange2:="X"+ALLTRIM(STR(INT(nRowIni)))+":Z"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Range(cRange2):Set( "NumberFormat", cMascMon)
      ENDIF
      cRange2:="A"+ALLTRIM(STR(INT(nRowIni)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range( cRange2 ):Borders():LineStyle  := 1
      //
      cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
      aResumen[1,1]:=IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      aResumen[2,1]:=nExporta
      nInternaNG:=nTotal2
      nPerGen:=nIvaGen-nRetGen
      nPerGenAdi:=nIvaGenAdi-nRetGenAdi
      nPerReduc:=nIvaReduc-nRetReduc
      aDet:={}
      IF oVen:nVForm=1
         cRange1:="M"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
         oHoja:Range( cRange1 ):Borders():LineStyle  := 1
         oHoja:Cells( nRow, 13):Value := "Totales"
         oHoja:Cells( nRow, 13 ):Font:Bold := .T.
         cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":N"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 14):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
         cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 15):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
         cRange2:="P"+ALLTRIM(STR(INT(nRowIni)))+":P"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 16):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
         cRange2:="Q"+ALLTRIM(STR(INT(nRowIni)))+":Q"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 17):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 17):Set( "NumberFormat", cMascMon )
         cRange2:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 18):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
         cRange2:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 20):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
         cRange2:="U"+ALLTRIM(STR(INT(nRowIni)))+":U"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 21):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 21):Set( "NumberFormat", cMascMon )
         cRange2:="V"+ALLTRIM(STR(INT(nRowIni)))+":V"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 22):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 22):Set( "NumberFormat", cMascMon )
         cRange2:="M"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
      ENDIF
      IF oVen:nVForm=3
         cRange1:="M"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
         oHoja:Range( cRange1 ):Borders():LineStyle  := 1
         oHoja:Cells( nRow, 13):Value := "Totales"
         oHoja:Cells( nRow, 13 ):Font:Bold := .T.
         cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":N"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 14):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
         cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 15):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
         cRange2:="P"+ALLTRIM(STR(INT(nRowIni)))+":P"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 16):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
         cRange2:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 18):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
         cRange2:="S"+ALLTRIM(STR(INT(nRowIni)))+":S"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 19):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 19):Set( "NumberFormat", cMascMon )
         cRange2:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 20):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
         cRange2:="M"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
      ENDIF
      IF oVen:nVForm=4
         cRange1:="M"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
         oHoja:Range( cRange1 ):Borders():LineStyle  := 1
         oHoja:Cells( nRow, 13):Value := "Totales"
         oHoja:Cells( nRow, 13 ):Font:Bold := .T.
         cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":N"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 14):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
         cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 15):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
         cRange2:="P"+ALLTRIM(STR(INT(nRowIni)))+":P"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 16):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
         cRange2:="Q"+ALLTRIM(STR(INT(nRowIni)))+":Q"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 17):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 17):Set( "NumberFormat", cMascMon )
         cRange2:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 18):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
         cRange2:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 20):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
         cRange2:="U"+ALLTRIM(STR(INT(nRowIni)))+":U"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 21):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 21):Set( "NumberFormat", cMascMon )
         cRange2:="V"+ALLTRIM(STR(INT(nRowIni)))+":V"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 22):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 22):Set( "NumberFormat", cMascMon )
         cRange2:="X"+ALLTRIM(STR(INT(nRowIni)))+":X"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 24):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 24):Set( "NumberFormat", cMascMon )
         cRange2:="Y"+ALLTRIM(STR(INT(nRowIni)))+":Y"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 25):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 25):Set( "NumberFormat", cMascMon )
         cRange2:="Z"+ALLTRIM(STR(INT(nRowIni)))+":Z"+ALLTRIM(STR(INT(nRowFin)))
         oHoja:Cells( nRow, 26):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja:Cells( nRow, 26):Set( "NumberFormat", cMascMon )
         cRange2:="M"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
      ENDIF
      IF oVen:lVSombTot
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
             oHoja:Cells( nRow, 14):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
             oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
             cRange2:=cLetra2+ALLTRIM(STR(INT(nRowIni2)))+":"+cLetra2+ALLTRIM(STR(INT(nRowFin)))
             oHoja:Cells( nRow, 15):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
             oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
             cRange2:=cLetra3+ALLTRIM(STR(INT(nRowIni2)))+":"+cLetra3+ALLTRIM(STR(INT(nRowFin)))
             oHoja:Cells( nRow, 18):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
             oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
             cRange2:=cLetra4+ALLTRIM(STR(INT(nRowIni2)))+":"+cLetra4+ALLTRIM(STR(INT(nRowFin)))
             oHoja:Cells( nRow, 20):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
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
      oHoja:Cells( nRow, 14):Value := "=+"+oVen:cEFunSum+"("+cRange1+")"
      oHoja:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
      IF oVen:lVSombTot
         oHoja:Cells( nRow, 14):Interior:Color := 12632256
      ENDIF
      cRange1:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 15):Value := "=+"+oVen:cEFunSum+"("+cRange1+")"
      oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
      IF oVen:lVSombTot
         oHoja:Cells( nRow, 15):Interior:Color := 12632256
      ENDIF
      cRange1:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 18):Value := "=+"+oVen:cEFunSum+"("+cRange1+")"
      oHoja:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
      IF oVen:lVSombTot
         oHoja:Cells( nRow, 18):Interior:Color := 12632256
      ENDIF
      cRange1:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 20):Value := "=+"+oVen:cEFunSum+"("+cRange1+")"
      oHoja:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
      IF oVen:lVSombTot
         oHoja:Cells( nRow, 20):Interior:Color := 12632256
      ENDIF
      nRow+=2
      oHojaX:=oHoja
      IF oVen:lVSepAju
         oHojaX:=oHoja2
         oHoja2:Cells:Font:Name := "Arial"
         oHoja2:Cells:Font:Size := 8
         oHoja2:PageSetup:Set("Zoom",oVen:nVPorAju)
         oHoja2:PageSetup:Set("Orientation",2 )
         oHoja2:PageSetup:Set("PaperSize" , 1 )
         oHoja2:PageSetup:Set("PrintTitleRows","A1:A"+ALLTRIM(STR(nDinamic)))
         IF oVen:lVNumPag
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
         IF oVen:nVForm=1 .OR. oVen:nVForm=4
            IF oVen:nVForm=1
               cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":T"+ALLTRIM(STR(INT(nRow+1)))
            ENDIF
            IF oVen:nVForm=4
               cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":X"+ALLTRIM(STR(INT(nRow+1)))
            ENDIF
            oHoja2:Range( cRange1 ):Borders():LineStyle  := 1
         ENDIF
         oHoja2:Cells(nRow,1):Value:="AJUSTES"
         oHoja2:Cells( nRow, 1 ):Font:Size := 10
         oHoja2:Cells( nRow, 1 ):Font:Bold := .T.
         IF oVen:nVForm=1 .OR. oVen:nVForm=4
            oHoja2:Cells(nRow,16):Value:="VENTAS INTERNAS O"
            oHoja2:Range( "P"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
            oHoja2:Cells( nRow, 16):Set("HorizontalAlignment",3)
            IF oVen:lVSombCab
               oHoja2:Range( cRange1 ):Interior:Color := 12632256
            ENDIF
            nRow++
            oHoja2:Cells(nRow,16):Value:="EXPORTACIONES GRAVADAS"
            oHoja2:Range( "P"+ALLTRIM(STR(nRow))+":T"+ALLTRIM(STR(nRow))):Merge()
            oHoja2:Cells( nRow, 16):Set("HorizontalAlignment",3)
            cRange1:="P"+ALLTRIM(STR(INT(nRow)))+":T"+ALLTRIM(STR(INT(nRow+1)))
            IF oVen:lVSombCab
               oHoja2:Range( cRange1 ):Interior:Color := 12632256
            ENDIF
         ENDIF
         nRow++
         nRowIni:=nRow
         nRowIniCab:=nRow
         nJ:=1
         FOR nI := 1 TO LEN(aCabecera)
             IF oVen:nVForm=1
                IF nI=16 .OR. nI=32
                   oHoja2:Range( "P"+ALLTRIM(STR(nRow))+":R"+ALLTRIM(STR(nRow))):Merge()
                   oHoja2:Cells( nRow, 16):Set("HorizontalAlignment",3)
                ENDIF
             ENDIF
             IF oVen:nVForm=4
                IF nI=16 .OR. nI=42
                   oHoja2:Range( "P"+ALLTRIM(STR(nRow))+":R"+ALLTRIM(STR(nRow))):Merge()
                   oHoja2:Cells( nRow, 16):Set("HorizontalAlignment",3)
                ENDIF
                IF nI=21 .OR. nI=47
                   oHoja2:Range( "U"+ALLTRIM(STR(nRow))+":X"+ALLTRIM(STR(nRow))):Merge()
                   oHoja2:Cells( nRow, 21):Set("HorizontalAlignment",3)
                ENDIF
             ENDIF
             oHoja2:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/3)) ):Value := aCabecera[nI]
             oHoja2:Cells( nRow, nI-((nJ-1)*(LEN(aCabecera)/3)) ):Set("HorizontalAlignment",3)
             IF LEN(aCabecera)/3=nI/nJ
                nRow++
                nJ++
             ENDIF
         NEXT
         IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
            cRange1:="A"+ALLTRIM(STR(INT(nRowIni)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRowIni+2)))
         ENDIF
         oHoja2:Range( cRange1 ):Borders():LineStyle  := 1
         IF oVen:lVSombCab
            oHoja2:Range( cRange1 ):Interior:Color := 12632256
         ENDIF
         nRowIni:=nRow
         nRowFin:=nRow
         // RETENCIONES NO DECLARADAS EN SU MOMENTO
         nI:=1
         cSql:="SELECT * FROM DPDOCCLIRTI WHERE RTI_FECHA"+GetWhere(">=",dFecha1)+" AND RTI_FECHA"+GetWhere("<=",dFecha2)+" ORDER BY RTI_FECHA"
         oTable:=OpenTable(cSql)
         DO WHILE lRetAju .AND. !oTable:Eof() .AND. oTable:RTI_FECHA<=dFecha2
            cSql:="SELECT "+;
                  "DOC_CODSUC, "+;
                  "DOC_TIPDOC, "+;
                  "DOC_NUMERO, "+;
                  "DOC_FECHA,  "+;
                  "DOC_ANUFIS, "+;
                  "DOC_NUMFIS, "+;
                  "CLI_CODIGO, "+;
                  "MOV_IVA   , "+;
                  "MOV_TIPIVA, "+;
                  "DOC_CXC   , "+;
                  "DOC_ACT   , "+;
                  "MOV_TOTAL , "+;
                  "MOV_IMPOTR, "+;
                  "DOC_DCTO  , "+;
                  "DOC_RECARG, "+;
                  "DOC_OTROS,  "+;
                  "DOC_NETO,   "+;
                  "IF(DOC_ANUFIS=0,TDC_LIBTRA,'03-Anu') AS TDC_LIBTRA, "+;
                  "IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_RIF   ,DPCLIENTES.CLI_RIF   ) AS CLI_RIF    ,"+;
                  "IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_NOMBRE,DPCLIENTES.CLI_NOMBRE) AS CLI_NOMBRE ,"+;
                  "IF(DOC_CODIGO='0000000000',DPCLIENTESCERO.CCG_CONTRI,DPCLIENTES.CLI_CONTRI) AS CLI_CONTRI  "+;
                  " FROM DPDOCCLI INNER JOIN DPCLIENTES ON CLI_CODIGO=DOC_CODIGO "+;
                  " INNER JOIN DPTIPDOCCLI ON DOC_TIPDOC=TDC_TIPO   "+;
                  " LEFT  JOIN DPCLIENTESCERO ON DOC_CODSUC=CCG_CODSUC AND "+;
                  "            DOC_TIPDOC=CCG_TIPDOC AND "+;
                  "            DOC_NUMERO=CCG_NUMDOC "+;
                  " LEFT  JOIN DPMOVINV ON MOV_CODSUC=DOC_CODSUC AND MOV_TIPDOC=DOC_TIPDOC AND MOV_CODCTA=DOC_CODIGO"+;
                  "            AND MOV_DOCUME=DOC_NUMERO AND MOV_INVACT=1 "+;
                  "WHERE DOC_CODSUC"+GetWhere("=",oTable:RTI_CODSUC)+" AND DOC_TIPDOC"+GetWhere("=",oTable:RTI_TIPDOC)+;
                  " AND DOC_NUMERO"+GetWhere("=",oTable:RTI_NUMERO)+" AND DOC_TIPTRA"+GetWhere("=",oTable:RTI_TIPTRA)
   
            oTableDoc:=OpenTable(cSql)
            IF oTableDoc:RecCount()=0 .OR. oTableDoc:DOC_ANUFIS .OR. LEFT(DTOS(oTableDoc:DOC_FECHA),6)>=LEFT(DTOS(oTable:RTI_FECHA),6)
               oTableDoc:End()
               oTable:Skip()
               LOOP
            ENDIF
            cEstado:=SQLGET("DPDOCCLI","DOC_ESTADO","DOC_CODSUC"+GetWhere("=",oTable:RTI_CODSUC)+" AND DOC_NUMERO"+GetWhere("=",oTable:RTI_DOCNUM)+" AND DOC_TIPDOC"+GetWhere("=",oTable:RTI_DOCTIP))
            IF cEstado="NU"
               oTableDoc:End()
               oTable:Skip()
               LOOP
            ENDIF
            oTableDoc:GoTop()
            oTableDoc:Replace("DOC_OPERAC",STRZERO(0,6)) // Número de Transacción
            oTableDoc:Replace("DOC_FACAFE",SPACE(10))    // Factura Afectada
            oTableDoc:Replace("DOC_MTOIVA",0        )    // Alicuota IVA
            oTableDoc:Replace("DOC_EXONER",0        )    // Exento
            oTableDoc:Replace("DOC_MTORTI",0        )    // Monto de Retención RTI
            lContri:= oTableDoc:CLI_CONTRI<>"N"
            cLetraBase:=IIF( !lContri,"R" ,"Q" )
            nBaseCol:=IIF( !lContri,18 ,17 )
            cLetraBas8:=IIF( !lContri,"V" ,"U" )
            nBaseCol8:=IIF( !lContri,22 ,21 )
            cRif:=PADR(oTableDoc:CLI_RIF,15)
            cRsocial:=PADR(oTableDoc:CLI_NOMBRE,oVen:nVLenRs)
            cDocNum:=oTable:RTI_NUMERO
            cDocFis:=PADR(oTableDoc:DOC_NUMFIS,10)
            cAfec:=PADR(oTableDoc:DOC_FACAFE,10)
            nTotDoc:=oTableDoc:DOC_NETO*oTableDoc:DOC_CXC*oTableDoc:DOC_ACT
            nBruto:=0
            IF Empty(oTableDoc:MOV_TIPIVA) // No tiene Productos
               aLine:=ACLONE(oTableDoc:aDataFill[oTableDoc:Recno()])
               EJECUTAR("DPDOCCLIIVA",oTableDoc:DOC_CODSUC,oTableDoc:DOC_TIPDOC,oTableDoc:CLI_CODIGO,oTableDoc:DOC_NUMERO,.F.,;
                                  oTableDoc:DOC_DCTO  ,oTableDoc:DOC_RECARG,oTableDoc:DOC_OTROS,NIL,"V")
               WHILE .T.
                 nAt:=ASCAN(oDp:aArrayIva,{|a,n|a[4]=0})
                 IF nAt=0
                   EXIT
                 ENDIF
                 oDp:aArrayIva:=ARREDUCE(oDp:aArrayIva,nAt)
               ENDDO
               FOR I:=1 TO LEN(oDp:aArrayIva)
                   nAt:=ASCAN(aTipIva,{|a,n|a[1]==oDp:aArrayIva[I,1]})
                   IF nAt>0 .AND. oDp:aArrayIva[I,4]<>0 .AND. LEN(oDp:aArrayIva)>0
                      aLine[oTableDoc:FieldPos("MOV_TIPIVA")]:=oDp:aArrayIva[I,1]
                      aLine[oTableDoc:FieldPos("MOV_TOTAL" )]:=oDp:aArrayIva[I,4]
                      aLine[oTableDoc:FieldPos("MOV_IVA"   )]:=oDp:aArrayIva[I,3]
                      IF I=1
                         oTableDoc:aDataFill[oTableDoc:Recno()]:=ACLONE(aLine)
                      ELSE
                         AADD(oTableDoc:aDataFill,ACLONE(aLine))
                      ENDIF
                   ENDIF
               NEXT I
            ENDIF
            nLine:=0
            nBase:=0
            nIva:=0
            nIvaGen:=0
            nIvaIcs:=0
            nIvaIcs2:=0
            nIvaRed:=0
            nImp15:=0
            nImp8:=0
            nPor8:=0
            nBase8:=0
            nBase15:=0
            oTableDoc:GoTop()
            WHILE !oTableDoc:Eof()
               IF oTableDoc:DOC_ANUFIS
                  oTableDoc:Replace("MOV_TOTAL",0)
               ENDIF
               oDp:nDesc     :=oTableDoc:DOC_DCTO*oTableDoc:DOC_ACT
               oDp:nRecarg   :=oTableDoc:DOC_RECARG*oTableDoc:DOC_ACT
               oDp:nDocOtros :=oTableDoc:DOC_OTROS*oTableDoc:DOC_ACT
               oDp:nBruto    :=oTableDoc:MOV_TOTAL*oTableDoc:DOC_ACT
               oDp:nMtoDesc  :=Porcen(oDp:nBruto,oDp:nDesc  )*oTableDoc:DOC_ACT
               oDp:nMtoRecarg:=Porcen(oDp:nBruto,oDp:nRecarg)*oTableDoc:DOC_ACT
               oDp:nNeto     :=oDp:nBruto+oDp:nMtoRecarg-oDp:nMtoDesc+oDp:nDocOtros
               oDp:nMtoVar   :=oDp:nBruto-oDp:nNeto
               oDp:nPorVar   :=(100-RATA(oDp:nNeto,oDp:nBruto))*-1
               oTableDoc:Replace("MOV_BASE"  ,(oTableDoc:MOV_TOTAL+PORCEN(oTableDoc:MOV_TOTAL,oDp:nPorVar))*oTableDoc:DOC_ACT)
               oTableDoc:Replace("DOC_MTOIVA",PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT)
               nLine:=nLine+1
               IF nLine=1 .AND. oTableDoc:MOV_IVA=0
                  oTableDoc:Replace("DOC_EXONER",oTableDoc:DOC_NETO*oTableDoc:DOC_ACT)
                  oTableDoc:Replace("MOV_BASE"  ,0) // No va para el Libro de Venta
               ENDIF
               IF nLine>1
                  oTableDoc:Replace("DOC_NETO",0)
                  nBaseDet :=oTableDoc:MOV_BASE*oTableDoc:DOC_ACT        // EXENTO
                  IF oTableDoc:MOV_IVA=0            // EXENTO
                     oTableDoc:Replace("MOV_BASE",0) // No va para el Libro de Venta
                     oTableDoc:Replace("DOC_NETO",0) //
                     nRecAct:=oTableDoc:Recno()
                     oTableDoc:Goto(nRecord)         // Registro Anterior
                     oTableDoc:Replace("DOC_EXONER",nBaseDet*oTableDoc:DOC_ACT)
                     oTableDoc:Goto(nRecAct)
                  ENDIF
               ENDIF
               nBase:=nBase+(oTableDoc:MOV_BASE*oTableDoc:DOC_ACT)
               nIva:=nIva+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               cSqlIva:="SELECT * FROM DPIVATABC WHERE CTI_FECHA"+GetWhere("<=",oTableDoc:DOC_FECHA)+" AND CTI_TIPO"+GetWhere("=",oTableDoc:MOV_TIPIVA)+" ORDER BY CTI_FECHA DESC LIMIT 1"
               oTableIva:=OpenTable(cSqlIva)
               IF oTableIva:RecCount()>0 .AND. oTableDoc:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="GN"
                  nIvaGen:=nIvaGen+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTableDoc:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="S1"
                  nIvaIcs:=nIvaIcs+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTableDoc:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="S2"
                  nIvaIcs2:=nIvaIcs2+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               ENDIF
               IF oTableIva:RecCount()>0 .AND. oTableDoc:MOV_TIPIVA=oTableIva:CTI_TIPO .AND. oTableIva:CTI_TIPO="RD"
                  nIvaRed:=nIvaRed+PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
                  nBase8:=nBase8+oTableDoc:MOV_BASE*oTableDoc:DOC_ACT
                  nPor8:=PORCEN(oTableDoc:MOV_BASE,oTableDoc:MOV_IVA)*oTableDoc:DOC_ACT
               ENDIF
               oTableDoc:Skip()
            ENDDO
            nImp15:=nIva+nIvaIcs+nIvaIcs2
            nImp8:=nIvaRed
            nBase15:=nBase-nBase8
            oTableDoc:GoTop()
            nExento:=oTableDoc:DOC_EXONER
            nAlic:=RATA(nIva,nBase)
            nIvaRet:=SQLGET("DPDOCCLI","DOC_NETO*DOC_ACT","DOC_CODSUC"+GetWhere("=",oTable:RTI_CODSUC)+" AND DOC_TIPDOC"+GetWhere("=",oTable:RTI_DOCTIP)+" AND DOC_NUMERO"+GetWhere("=",oTable:RTI_DOCNUM)+" AND DOC_TIPTRA='D'")
            nFactor:=oTableDoc:DOC_CXC
            nTotDoc:=nTotDoc*nFactor
            nBase:=nBase*nFactor
            nExento:=nExento*nFactor
            nIva:=nIva*nFactor
            nIvaRet:=nIvaRet*nFactor
            nBase8:=nBase8*nFactor
            nBase15:=nBase15*nFactor
            nImp8:=nImp8*nFactor
            nImp15:=nImp15*nFactor
            nIvaPer:=nIva-nIvaRet
            cCadena1:=SPACE(54)
            // DR20110201a. Se agrega condición para FCT
            IF oTableDoc:DOC_TIPDOC="FAV"  .OR. oTableDoc:DOC_TIPDOC="FCT"
               cCadena1:=PADC(cDocNum,10)+" "+PADC(cDocFis,10)+" "+SPACE(10)+" "+SPACE(10)+" "+PADC(oTableDoc:TDC_LIBTRA,10)
            ENDIF
            IF oTableDoc:DOC_TIPDOC="DEB"
               cCadena1:=SPACE(10)+" "+PADC(cDocFis,10)+" "+PADC(cDocNum,10)+" "+SPACE(10)+" "+PADC(oTableDoc:TDC_LIBTRA,10)
            ENDIF
            IF oTableDoc:DOC_TIPDOC="CRE" .OR. oTableDoc:DOC_TIPDOC="DEV"
               cCadena1:=SPACE(10)+" "+PADC(cDocFis,10)+" "+SPACE(10)+" "+PADC(cDocNum,10)+" "+PADC(oTableDoc:TDC_LIBTRA,10)
            ENDIF
            cCadena2:=cAfec+" "+PADL(ALLTRIM(TRANSFORM(nTotDoc,"999,999,999.99")),14)
            cCadena3:=PADL(ALLTRIM(TRANSFORM(nExento,"999,999,999.99")),14)+" "+PADL(ALLTRIM(TRANSFORM(nBase,"999,999,999.99")),14)
            cCadena4:=PADL(ALLTRIM(TRANSFORM(nIvaRet,"999,999,999.99")),14)+" "+PADL(ALLTRIM(TRANSFORM(nIvaPer,"999,999,999.99")),14)
            oHoja2:Cells( nRow, 01):Value := IIF( oVen:lVConDin,'='+oVen:cEFunCel+'("ROW"'+oVen:cESepLis+'A'+ALLTRIM(STR(nRow))+')-'+ALLTRIM(STR(nDinamic)),nI)
            oHoja2:Cells( nRow, 02):Value := oTableDoc:DOC_FECHA
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
            IF oVen:nVForm=1 .OR. oVen:nVForm=4
               IF oVen:nVForm=1
                  oHoja2:Cells( nRow, nBaseCol):Value := nBase
                  oHoja2:Cells( nRow, nBaseCol):Set( "NumberFormat", cMascMon )
                  oHoja2:Cells( nRow, 19):Value := nAlic/100
                  oHoja2:Cells( nRow, 19):Set( "NumberFormat", cMascPor )
                  oHoja2:Cells( nRow, 20):Value := nIva
                  oHoja2:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
               ENDIF
               IF oVen:nVForm=4
                  oHoja2:Cells( nRow, nBaseCol):Value := nBase15
                  oHoja2:Cells( nRow, nBaseCol):Set( "NumberFormat", cMascMon )
                  oHoja2:Cells( nRow, 19):Value := RATA(nImp15,nBase15)/100
                  oHoja2:Cells( nRow, 19):Set( "NumberFormat", cMascPor )
                  oHoja2:Cells( nRow, 20):Value := nImp15
                  oHoja2:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
                  oHoja2:Cells( nRow, nBaseCol8):Value := nBase8
                  oHoja2:Cells( nRow, nBaseCol8):Set( "NumberFormat", cMascMon )
                  oHoja2:Cells( nRow, 23):Value := Rata(nImp8,nBase8)/100
                  oHoja2:Cells( nRow, 23):Set( "NumberFormat", cMascPor )
                  oHoja2:Cells( nRow, 24):Value := nImp8
                  oHoja2:Cells( nRow, 24):Set( "NumberFormat", cMascMon )
               ENDIF
               oHoja2:Cells( nRow, 21):Value := nIvaRet
               oHoja2:Cells( nRow, 21):Set( "NumberFormat", cMascMon )
            ENDIF
            IF oVen:nVForm=3
               oHoja2:Cells( nRow, 16):Value := nBase
               oHoja2:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
               oHoja2:Cells( nRow, 17):Value := nAlic/100
               oHoja2:Cells( nRow, 17):Set( "NumberFormat", cMascPor )
               oHoja2:Cells( nRow, 18):Value := nIva
               oHoja2:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
               oHoja2:Cells( nRow, 19):Value := nIvaRet
               oHoja2:Cells( nRow, 19):Set( "NumberFormat", cMascMon )
            ENDIF
            nRowFin:=nRow
            IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
               cRange1:="A"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
            ENDIF
            oHoja2:Range( cRange1 ):Borders():LineStyle  := 1
            nRow++
            nI++
            oTableDoc:End()
            oTable:Skip()
         ENDDO
         //
         nRow++
         IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
            cRange1:="M"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
         ENDIF
         oHoja2:Range( cRange1 ):Borders():LineStyle  := 1
         oHoja2:Cells( nRow, 13):Value := "Totales"
         oHoja2:Cells( nRow, 13 ):Font:Bold := .T.
         cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":N"+ALLTRIM(STR(INT(nRowFin)))
         oHoja2:Cells( nRow, 14):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja2:Cells( nRow, 14):Set( "NumberFormat", cMascMon )
         cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
         oHoja2:Cells( nRow, 15):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja2:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
         cRange2:="P"+ALLTRIM(STR(INT(nRowIni)))+":P"+ALLTRIM(STR(INT(nRowFin)))
         oHoja2:Cells( nRow, 16):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja2:Cells( nRow, 16):Set( "NumberFormat", cMascMon )
         IF oVen:nVForm=1 .OR. oVen:nVForm=4
            cRange2:="Q"+ALLTRIM(STR(INT(nRowIni)))+":Q"+ALLTRIM(STR(INT(nRowFin)))
            oHoja2:Cells( nRow, 17):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
            oHoja2:Cells( nRow, 17):Set( "NumberFormat", cMascMon )
         ENDIF
         cRange2:="R"+ALLTRIM(STR(INT(nRowIni)))+":R"+ALLTRIM(STR(INT(nRowFin)))
         oHoja2:Cells( nRow, 18):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja2:Cells( nRow, 18):Set( "NumberFormat", cMascMon )
         IF oVen:nVForm=3
            cRange2:="S"+ALLTRIM(STR(INT(nRowIni)))+":S"+ALLTRIM(STR(INT(nRowFin)))
            oHoja2:Cells( nRow, 19):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
            oHoja2:Cells( nRow, 19):Set( "NumberFormat", cMascMon )
         ENDIF
         cRange2:="T"+ALLTRIM(STR(INT(nRowIni)))+":T"+ALLTRIM(STR(INT(nRowFin)))
         oHoja2:Cells( nRow, 20):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
         oHoja2:Cells( nRow, 20):Set( "NumberFormat", cMascMon )
         IF oVen:nVForm=1 .OR. oVen:nVForm=4
            cRange2:="U"+ALLTRIM(STR(INT(nRowIni)))+":U"+ALLTRIM(STR(INT(nRowFin)))
            oHoja2:Cells( nRow, 21):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
            oHoja2:Cells( nRow, 21):Set( "NumberFormat", cMascMon )
            cRange2:="V"+ALLTRIM(STR(INT(nRowIni)))+":V"+ALLTRIM(STR(INT(nRowFin)))
            oHoja2:Cells( nRow, 22):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
            oHoja2:Cells( nRow, 22):Set( "NumberFormat", cMascMon )
         ENDIF
         IF oVen:nVForm=4
            cRange2:="X"+ALLTRIM(STR(INT(nRowIni)))+":X"+ALLTRIM(STR(INT(nRowFin)))
            oHoja2:Cells( nRow, 24):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
            oHoja2:Cells( nRow, 24):Set( "NumberFormat", cMascMon )
            cRange2:="Y"+ALLTRIM(STR(INT(nRowIni)))+":Y"+ALLTRIM(STR(INT(nRowFin)))
            oHoja2:Cells( nRow, 25):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
            oHoja2:Cells( nRow, 25):Set( "NumberFormat", cMascMon )
            cRange2:="Z"+ALLTRIM(STR(INT(nRowIni)))+":Z"+ALLTRIM(STR(INT(nRowFin)))
            oHoja2:Cells( nRow, 26):Value := IIF( .F. ,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
            oHoja2:Cells( nRow, 26):Set( "NumberFormat", cMascMon )
         ENDIF
         IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
            cRange2:="M"+ALLTRIM(STR(INT(nRow)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRow)))
         ENDIF

         IF oVen:lVSombTot
            oHoja2:Range( cRange2 ):Interior:Color := 12632256
         ENDIF
         nRow+=2
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
         IF oVen:lVSombTot
            oHojaX:Range( cRange1 ):Interior:Color := 12632256
         ENDIF
      ELSE
         AADD(aDet,SPACE(7))
         AADD(aDet,SPACE(11))
         AADD(aDet,SPACE(15))
         AADD(aDet,SPACE(oVen:nVLenRs))
         AADD(aDet,SPACE(1))
         AADD(aDet,SPACE(15))
         AADD(aDet,SPACE(1))
         AADD(aDet,SPACE(10))
         AADD(aDet,SPACE(10))
         AADD(aDet,SPACE(10))
         AADD(aDet,SPACE(10))
         AADD(aDet,SPACE(10))
         AADD(aDet,SPACE(10))
         AADD(aDet,PADL(ALLTRIM(TRANSFORM(nTotal1,"999,999,999.99")),14)+" ")
         AADD(aDet,PADL(ALLTRIM(TRANSFORM(nTotal2,"999,999,999.99")),14)+" ")
         AADD(aDet,PADL(ALLTRIM(TRANSFORM(nTotal3,"999,999,999.99")),14)+" ")
         AADD(aDet,SPACE(7))
         AADD(aDet,PADL(ALLTRIM(TRANSFORM(nTotal4,"999,999,999.99")),14)+" ")
         AADD(aDet,PADL(ALLTRIM(TRANSFORM(nTotal5,"999,999,999.99")),14)+" ")
         AADD(aDet,PADL(ALLTRIM(TRANSFORM(nTotal6,"999,999,999.99")),14)+" ")
      ENDIF
   ENDIF
   IF oVen:nVForm=2
      cRange2:="B"+ALLTRIM(STR(INT(nRowIni)))+":B"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", "@")
      cRange2:="E"+ALLTRIM(STR(INT(nRowIni)))+":G"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", cMascMon)
      cRange2:="I"+ALLTRIM(STR(INT(nRowIni)))+":J"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", cMascMon)
      cRange2:="L"+ALLTRIM(STR(INT(nRowIni)))+":M"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", cMascMon)
      cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", cMascMon)
      cRange2:="H"+ALLTRIM(STR(INT(nRowIni)))+":H"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", cMascPor)
      cRange2:="K"+ALLTRIM(STR(INT(nRowIni)))+":K"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", cMascPor)
      cRange2:="N"+ALLTRIM(STR(INT(nRowIni)))+":N"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range(cRange2):Set( "NumberFormat", cMascPor)
      cRange2:="A"+ALLTRIM(STR(INT(nRowIni)))+":"+aAbc[nUltCol]+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Range( cRange2 ):Borders():LineStyle  := 1
      //
   *  cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
      cRange1:="D"+ALLTRIM(STR(INT(nRow)))+":O"+ALLTRIM(STR(INT(nRow)))
      oHoja:Range( cRange1 ):Borders():LineStyle  := 1
      oHoja:Cells( nRow, 04):Value := "Totales"
      oHoja:Cells( nRow, 04 ):Font:Bold := .T.
      cRange2:="E"+ALLTRIM(STR(INT(nRowIni)))+":E"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 05):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      oHoja:Cells( nRow, 05):Set( "NumberFormat", cMascMon )
      cRange2:="F"+ALLTRIM(STR(INT(nRowIni)))+":F"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 06):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      oHoja:Cells( nRow, 06):Set( "NumberFormat", cMascMon )
      cRange2:="G"+ALLTRIM(STR(INT(nRowIni)))+":G"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 07):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      oHoja:Cells( nRow, 07):Set( "NumberFormat", cMascMon )
      cRange2:="I"+ALLTRIM(STR(INT(nRowIni)))+":I"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 09):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      oHoja:Cells( nRow, 09):Set( "NumberFormat", cMascMon )
      cRange2:="J"+ALLTRIM(STR(INT(nRowIni)))+":J"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 10):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      oHoja:Cells( nRow, 10):Set( "NumberFormat", cMascMon )
      cRange2:="L"+ALLTRIM(STR(INT(nRowIni)))+":L"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 12):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      oHoja:Cells( nRow, 12):Set( "NumberFormat", cMascMon )
      cRange2:="M"+ALLTRIM(STR(INT(nRowIni)))+":M"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 13):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      oHoja:Cells( nRow, 13):Set( "NumberFormat", cMascMon )
      cRange2:="O"+ALLTRIM(STR(INT(nRowIni)))+":O"+ALLTRIM(STR(INT(nRowFin)))
      oHoja:Cells( nRow, 15):Value := IIF( lPrimera,0 ,"=+"+oVen:cEFunSum+"("+cRange2+")")
      oHoja:Cells( nRow, 15):Set( "NumberFormat", cMascMon )
      cRange2:="D"+ALLTRIM(STR(INT(nRow)))+":O"+ALLTRIM(STR(INT(nRow)))
      IF oVen:lVSombTot
         oHoja:Range( cRange2 ):Interior:Color := 12632256
      ENDIF
   ENDIF
   IF oVen:nVForm=1 .OR. oVen:nVForm=3 .OR. oVen:nVForm=4
      FOR nI := 1 TO 20
          IF oVen:nVForm=1
             IF (nI>1 .AND. nI<8) .OR. (nI>13 .AND. nI<19) .OR. nI>19
                oHoja:Columns( nI ):AutoFit()
             ENDIF
          ENDIF
          IF oVen:nVForm=3
             IF (nI>1 .AND. nI<8) .OR. (nI>13 .AND. nI<17) .OR. nI>17
                oHoja:Columns( nI ):AutoFit()
             ENDIF
          ENDIF
          IF oVen:nVForm=4
             IF (nI>1 .AND. nI<8) .OR. (nI>13 .AND. nI<19) .OR. (nI>19 .AND. nI<23) .OR. nI>23
                oHoja:Columns( nI ):AutoFit()
             ENDIF
          ENDIF
      NEXT
      oHoja:Columns( 01 ):Set("ColumnWidth",3.14)
      oHoja:Columns( 08 ):Set("ColumnWidth",8.71)
      oHoja:Columns( 09 ):Set("ColumnWidth",8.71)
      oHoja:Columns( 10 ):Set("ColumnWidth",8.71)
      oHoja:Columns( 11 ):Set("ColumnWidth",8.71)
      oHoja:Columns( 12 ):Set("ColumnWidth",5.57)
      oHoja:Columns( 13 ):Set("ColumnWidth",8.71)
      IF oVen:nVForm=1 .OR. oVen:nVForm=4
         oHoja:Columns( 19 ):Set("ColumnWidth",5.57)
      ENDIF
      IF oVen:nVForm=3
         oHoja:Columns( 17 ):Set("ColumnWidth",5.57)
      ENDIF
      IF oVen:nVForm=4
         oHoja:Columns( 23 ):Set("ColumnWidth",5.57)
      ENDIF
      IF oVen:lVSepAju
         FOR nI := 1 TO 20
             IF oVen:nVForm=1
                IF (nI>1 .AND. nI<8) .OR. (nI>13 .AND. nI<19) .OR. nI>19
                   oHoja2:Columns( nI ):AutoFit()
                ENDIF
             ENDIF
             IF oVen:nVForm=3
                IF (nI>1 .AND. nI<8) .OR. (nI>13 .AND. nI<17) .OR. nI>17
                   oHoja2:Columns( nI ):AutoFit()
                ENDIF
                IF (nI>1 .AND. nI<8) .OR. (nI>13 .AND. nI<19) .OR. (nI>19 .AND. nI<23) .OR. nI>23
                   oHoja2:Columns( nI ):AutoFit()
                ENDIF
             ENDIF
         NEXT
         oHoja2:Columns( 01 ):Set("ColumnWidth",3.14)
         oHoja2:Columns( 08 ):Set("ColumnWidth",8.71)
         oHoja2:Columns( 09 ):Set("ColumnWidth",8.71)
         oHoja2:Columns( 10 ):Set("ColumnWidth",8.71)
         oHoja2:Columns( 11 ):Set("ColumnWidth",8.71)
         oHoja2:Columns( 12 ):Set("ColumnWidth",5.57)
         oHoja2:Columns( 13 ):Set("ColumnWidth",8.71)
         IF oVen:nVForm=1 .OR. oVen:nVForm=4
            oHoja2:Columns( 19 ):Set("ColumnWidth",5.57)
            oHoja:Columns( 23 ):Set("ColumnWidth",0.01)
         ENDIF
         IF oVen:nVForm=3
            oHoja2:Columns( 17 ):Set("ColumnWidth",5.57)
            oHoja:Columns( 21 ):Set("ColumnWidth",0.01)
         ENDIF
         IF oVen:nVForm=4
            oHoja2:Columns( 23 ):Set("ColumnWidth",5.57)
            oHoja:Columns( 27 ):Set("ColumnWidth",0.01)
         ENDIF
         *oHoja2:Cells( 1, 1 ):Select()
      ENDIF
   ENDIF
   IF oVen:nVForm=2
      oHoja:Columns( 01 ):Set("ColumnWidth",8.00)
      oHoja:Columns( 02 ):AutoFit()
      oHoja:Columns( 03 ):Set("ColumnWidth",25.00)
      oHoja:Columns( 04 ):AutoFit()
      oHoja:Columns( 05 ):Set("ColumnWidth",13.00)
      oHoja:Columns( 06 ):Set("ColumnWidth",13.00)
      oHoja:Columns( 07 ):Set("ColumnWidth",12.29)
      oHoja:Columns( 08 ):Set("ColumnWidth",5.71)
      oHoja:Columns( 09 ):Set("ColumnWidth",12.29)
      oHoja:Columns( 10 ):Set("ColumnWidth",12.29)
      oHoja:Columns( 11 ):Set("ColumnWidth",5.71)
      oHoja:Columns( 12 ):Set("ColumnWidth",12.29)
      oHoja:Columns( 13 ):Set("ColumnWidth",12.29)
      oHoja:Columns( 14 ):Set("ColumnWidth",5.71)
      oHoja:Columns( 15 ):Set("ColumnWidth",12.29)
      oHoja:Columns( 16 ):Set("ColumnWidth",0.01)
      IF oVen:lVColRti
         oHoja:Columns( 16 ):AutoFit()
         oHoja:Columns( 17 ):Set("ColumnWidth",0.01)
      ENDIF
   ENDIF
   oExcel:Visible := .T.

   oHoja:Activate()
   oHoja:End()
   oHoja:=NIL
   oHojaSuma:End()
   oHojaSuma:=NIL
   IF oVen:lVSepAju
      oHoja2:End()
      oHoja2:=NIL
   ENDIF
   oExcel:End()
   oExcel:=NIL
RETURN .T.
