/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código BALACECOMPROBACION
// Fecha  30/03/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodCta,cDescri,nAt,cWhere,nMonto
    LOCAL oTable,oXls,oCta
    LOCAL nLinIni,nContar,I,U,cItem,nItem:=0,cNumero:=STRZERO(1,8),cNumEje:=EJECUTAR("GETNUMEJE",oDp:dFchInicio)

    DEFAULT cCodigo:="BALACECOMPROBACION"

    oTable:=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    oTable:End(.T.)

    cWhere:="MOC_ORIGEN"+GetWhere("=","BAL")

    IF COUNT(cTable,cWhere)>0 .AND. MsgYesNo("Desea Remover todos los Asientos del Balacen Inicial ")
      SQLDELETE(cTable)
    ENDIF

    cWhere:="CBT_CODSUC"+GetWhere("=",oDp:cSucursal     )+" AND "+;
            "CBT_ACTUAL"+GetWhere("=","S"               )+" AND "+;
            "CBT_NUMERO"+GetWhere("=",cNumero           )+" AND "+;
            "CBT_FECHA" +GetWhere("=",oDp:dFchInicio    )

    IF !ISSQLFIND("DPCBTE",cWhere)
 
       EJECUTAR("CREATERECORD","DPCBTE",{"CBT_CODSUC" ,"CBT_ACTUAL","CBT_NUMERO","CBT_FECHA"   ,"CBT_NUMEJE","CBT_COMEN1"     },;
                                        {oDp:cSucursal,"S"         ,cNumero     ,oDp:dFchInicio,cNumEje     ,"Balance Inicial"},;
       NIL,.T.,cWhere)

    ENDIF

    SET DECI TO 2

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    oCta  :=OpenTable("SELECT * FROM DPCTA"  , .F. )
    oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
    oTable:lAuditar:=.F.
    oTable:SetForeignkeyOff()

// oXls:Browse()

    oXls:Gotop()

    WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cCodCta:=CTOO(oXls:COL_A,"C")

      IF "TOTAL"$UPPER(cCodCta)
         oXls:DbSkip()
         LOOP
      ENDIF

      nAt    :=AT("-",cCodCta)
      cDescri:=SUBS(cCodCta,nAt+1,LEN(cCodCta))
      cCodCta:=LEFT(cCodCta,nAt-1)
     
      // cCodCta:=STRTRAN(ALLTRIM(oXls:COL_A),"-","")
      cCodCta:=STRTRAN(cCodCta,";","")
      cCodCta:=STRTRAN(cCodCta,"/","")
      cCodCta:=STRTRAN(cCodCta,".","")
      cCodCta:=ALLTRIM(cCodCta)

      IF Empty(cCodCta) .OR. Empty(cDescri)
         oXls:DbSkip()
         LOOP
      ENDIF

      nMonto:=CTOO(oXls:COL_F,"N")

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodCta+CRLF),NIL)

      cWhere    :="CTA_CODMOD"+GetWhere("=",oDp:cCtaMod)+" AND CTA_CODIGO"+GetWhere("=",cCodCta)

      IF !ISSQLFIND("DPCTA",cWhere)
        oCta:AppendBlank()
        oCta:Replace("CTA_CODIGO",cCodCta)
        oCta:Replace("CTA_DESCRI",cDescri)
        oCta:Replace("CTA_CODMOD",oDp:cCtaMod)
        oCta:Commit("")
      ENDIF

      IF nMonto<>0

         nItem++ 
         cItem:=STRZERO(nItem,4)
         oTable:AppendBlank()
         oTable:Replace("MOC_ITEM"  ,cItem)
         oTable:Replace("MOC_CUENTA",cCodCta)
         oTable:Replace("MOC_DESCRI","Balance Inicial")
         oTable:Replace("MOC_CODMOD",oDp:cCtaMod)
         oTable:Replace("MOC_FECHA" ,oDp:dFchInicio)
         oTable:Replace("MOC_ACTUAL","S")
         oTable:Replace("MOC_ORIGEN","BAL")
         oTable:Replace("MOC_NUMCBT",cNumero)
         oTable:Replace("MOC_USUARI",oDp:cUsuario)
         oTable:Replace("MOC_MONTO" ,nMonto)

         oTable:Commit("")

      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Asignando Cuentas de Movimiento"+CRLF),NIL)
   EJECUTAR("SETDPCTADET")
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oCta:End()
   oXls:End()

   EJECUTAR("DPCTAUSO")


RETURN .T.

