/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código COLEGIO_CAPACHO
// Fecha  25/01/2025
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cWhere,cSql,dFchAnt:=CTOD("")
    LOCAL oTable,oXls,oDb:=OpenOdbc(oDp:cDsnData),cDescri:=""
    LOCAL nLinIni,nContar,I,U,nItem:=0
    LOCAL aRef:={}
    LOCAL cItem,cNumEje,nValCam:=1,cNumero:=STRZERO(1,8),dFecha,nMonto:=0,aFechas:={},cCodCta

    DEFAULT cCodigo:="COLEGIO_CAPACHO_NOM"

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_MEMO FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    aRef    :=EJECUTAR("IXLLOAD",NIL,NIL,NIL,NIL,oTable:IXL_MEMO)
    AEVAL(aRef,{|a,n| aRef[n,1]:=STRTRAN(a[1],"@","")})
    oTable:End(.T.)

    cDescri:=cFileNoPath(cFileXls)

    cWhere:="CBT_COMEN2"+GetWhere("=",cDescri)


    IF COUNT("DPCBTE",cWhere)>0 .AND. MsgYesNo("Desea Remover todos Comprobantes Importados desde XLS"+cDescri)

      cSql=[ DELETE FROM dpasientos ]+;
           [ INNER JOIN dpcbte ON MOC_CODSUC=CBT_CODSUC AND MOC_ACTUAL=CBT_ACTUAL AND MOC_FECHA=CBT_FECHA AND MOC_NUMCBT=CBT_NUMERO  ]+;
           [ WHERE ]+cWhere

      oDb:EXECUTE(cSql)

      SQLDELETE("DPCBTE",cWhere) // debe remover en cascada los asientos

    ENDIF

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

    IF(oImpXls:lBrowse,oXls:Browse(),NIL)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
    IF Empty(oTable:aDefault)
       EJECUTAR("DPTABLESETDEF",oTable) // Facilita Asinar valores por defecto  en :AppendBlank()
    ENDIF

    SET DECI TO 2

    oTable:lAuditar:=.F.
    oTable:SetForeignkeyOff()

    oXls:Gotop()

    WHILE !oXls:Eof() 
  
      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      // Asigna los valores desde oXls->oRef oReg:RIF = oXls:A
      AEVAL(aRef,{|a,n| aRef[n,2]:=MacroEje("oXls:COL_"+a[3]),;
                        oRef:Set(aRef[n,1],aRef[n,2])})

      cCodCta:=CTOO(oXls:COL_B,"C")
      cCodCta:=STRTRAN(cCodCta,"-","")
      cCodCta:=STRTRAN(cCodCta,";","")
      cCodCta:=STRTRAN(cCodCta,"/","")

      IF !Empty(CTOO(oXls:COL_A,"D"))
         dFecha:=CTOO(oXls:COL_A,"D")
      ENDIF

      IF Empty(cCodCta) .OR. Empty(dFecha)
         oXls:DbSkip()
         LOOP
      ENDIF

      cNumero:=EJECUTAR("DPNUMCBTEXTIPDOC","DPCBTE","CBT",dFecha)


//     
//      IF YEAR(dFecha)<2024
//         dFecha:=CTOD("31/01/2024")
//      ENDIF

      IF ValType(oXls:COL_A)="D"
         dFecha:=oXls:COL_A
      ENDIF

      dFecha :=FCHFINMES(dFecha) // solicitado por Lcdo Chona 24/01/2025

      cNumEje:=EJECUTAR("GETNUMEJE",dFecha)
      nValCam:=EJECUTAR("DPGETVALCAM",oDp:cMonedaExt,dFecha)

      cWhere:="CBT_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
              "CBT_ACTUAL"+GetWhere("=","S"          )+" AND "+;
              "CBT_NUMERO"+GetWhere("=",cNumero      )+" AND "+;
              "CBT_FECHA" +GetWhere("=",dFecha       )

      IF dFecha<>dFchAnt .AND. !ISSQLFIND("DPCBTE",cWhere)
 
        EJECUTAR("CREATERECORD","DPCBTE",{"CBT_CODSUC" ,"CBT_ACTUAL","CBT_NUMERO","CBT_FECHA"   ,"CBT_NUMEJE","CBT_COMEN2","CBT_ORIGEN"     },;
                                         {oDp:cSucursal,"S"         ,cNumero     ,dFecha        ,cNumEje     ,cDescri     ,"XLS"       },;
        NIL,.T.,cWhere)

      ENDIF

      IF dFecha<>dFchAnt
         nItem  :=0
         dFchAnt:=dFecha
      ENDIF

      cItem:=STRZERO(nItem,4)

      // oXls:COL_E:=ALLTRIM(CTOO(oXls:COL_E,"C")) // REFERENCIA
      oXls:COL_C:=CTOO(oXls:COL_C,"C")

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodCta+CRLF),NIL)

      nMonto :=CTOO(oXls:COL_F,"N")-CTOO(oXls:COL_G,"N") 
      nMtoIva:=0 //

      IF nMonto<>0

         nItem++ 
         cItem:=STRZERO(nItem,4)
         oTable:AppendBlank()
         oTable:Replace("MOC_ITEM"  ,cItem)
         oTable:Replace("MOC_CUENTA",cCodCta)
         oTable:Replace("MOC_DESCRI",oXls:COL_C)
         oTable:Replace("MOC_CODMOD",oDp:cCtaMod)
         oTable:Replace("MOC_DOCUME",""    ) // oXls:COL_B)
         oTable:Replace("MOC_FECHA" ,dFecha) // oDp:dFchInicio)
         oTable:Replace("MOC_ACTUAL","S")
         oTable:Replace("MOC_ORIGEN","XLS")
         oTable:Replace("MOC_NUMCBT",cNumero)
         oTable:Replace("MOC_NUMEJE",cNumEje)
         oTable:Replace("MOC_USUARI",oDp:cUsuario)
         oTable:Replace("MOC_MONTO" ,nMonto)
         oTable:Replace("MOC_VALCAM",nValCam)
//       oTable:Replace("MOC_RIF"   ,oXls:COL_C)
         oTable:Replace("MOC_CODSUC",oDp:cSucursal)
         oTable:Commit("")

      ENDIF

      oXls:DbSkip()

      SysRefresh(.T.)

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

   IF Type("oBALINIDIV")="O" .AND. oBALINIDIV:oWnd:hWnd>0
      oBALINIDIV:CLOSE()
   ENDIF

   IF(ValType(oMemo )="O",oMemo:Append("Creando Comprobante Inicial "+CRLF),NIL)
   EJECUTAR("DPCBTEFIX")

   EJECUTAR("BRASIENTORESORG","MOC_ORIGEN"+GetWhere("=","XLS"))

RETURN .T.
