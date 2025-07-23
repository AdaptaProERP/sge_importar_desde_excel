/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código BCO_0134
// Fecha  08/04/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere
    LOCAL oTable,oXls
    LOCAL nLinIni,nContar,I,U
    LOCAL cCodSuc:="",nItem:=0,cTitle:=NIL
    LOCAL dDesde :=CTOD("")
    LOCAL dHasta :=CTOD("")

    DEFAULT cCodigo:="BCO_0134",;
            oParXls:cCtaBanco:=""

    IF ValType(oParXls:cCtaBanco)="C" .AND. Empty(oParXls:cCtaBanco)
       MsgMemo("Necesario Seleccionar la cuenta Bancaria")
       RETURN .F.
    ENDIF

    IF Empty(oParXls:cCodBco) .AND. !EJECUTAR("DPREGLEEEDOCTAVALID",cCodigo,oMemo)
       RETURN .F.
    ENDIF

    cCodSuc:=SQLGET("DPCTABANCO","BCO_CODSUC","BCO_CTABAN"+GetWhere("=",oParXls:cCtaBanco))

    IF(ValType(oMemo)="O",oMemo:Append("Cuenta "+oParXls:cCtaBanco +CRLF),NIL)
    IF(ValType(oMemo)="O",oMemo:Append("Cuenta "+oParXls:cBcoNombre+CRLF),NIL)

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    oTable:End(.T.)

    cWhere :="RLE_CUENTA"+GetWhere("=",oParXls:cCodCta)+" AND "+;
             "RLE_FILE"  +GetWhere("=",cFileXls)

    // oParXls:cNumero:=SQLINCREMENTAL("DPREGLEEEDOCTA","RLE_NUMERO",cWhere)

    IF COUNT("DPREGLEEEDOCTA",cWhere)>0 .AND. MsgYesNo("Desea Remover los Registros de la tabla "+cTable+CRLF+"Asociados con "+cFileXls)
      SQLDELETE(cTable)
    ENDIF

    // oParXls:cNumero:=SQLINCREMENTAL("DPREGLEEEDOCTA","RLE_NUMERO",cWhere)

    IF(ValType(oSay)="O",oSay:SetText("Creando Registro"),NIL)

    EJECUTAR("DPREGLEEEDOCTACREA",cFileXls)

    IF(ValType(oSay)="O",oSay:SetText("Registro "+oParXls:cNumero),NIL)
    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    // Parámetros Bancarios
    oParXls:cCodBco  
    oParXls:cCtaBanco
    oParXls:cCodIBP  
    oParXls:cNumero  

/*
    cWhere :="RLE_CODBCO"+GetWhere("=",oParXls:cCodBco)+" AND "+;
             "RLE_CUENTA"+GetWhere("=",oParXls:cCodCta)+" AND "+;
             "RLE_NUMERO"+GetWhere("=",oParXls:cNumero)

    oParXls:cNumero:=SQLINCREMENTAL("DPREGLEEEDOCTA","RLE_NUMERO",cWhere)
*/
    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

// ViewArray(oXls:aData)
// oXls:Browse()
// 
// RETURN 

   IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

   oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
   oTable:lAuditar:=.F.
   oTable:SetForeignkeyOff()

   oXls:Gotop()

   WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      nItem++

/*
      cCodigo:=STRTRAN(ALLTRIM(oXls:COL_A),"-","")
      cCodigo:=STRTRAN(cCodigo,";","")
      cCodigo:=STRTRAN(cCodigo,"/","")

      IF Empty(cCodigo)
         oXls:DbSkip()
         LOOP
      ENDIF
*/

      oXls:COL_A:=CTOO(oXls:COL_A,"D")
      oXls:COL_D:=CTOO(oXls:COL_D,"N")
      oXls:COL_E:=CTOO(oXls:COL_E,"N")
      oXls:COL_B:=CTOO(oXls:COL_B,"C")
      

      cCodigo:=CTOO(oXls:COL_A,"C")+"-"+CTOO(oXls:COL_B,"C")

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodigo+CRLF),NIL)

      cWhere    :="ECB_FECHA" +GetWhere("=",oXls:COL_A)+" AND "+;
                  "ECB_REFERE"+GetWhere("=",oXls:COL_B)+" AND "+;
                  "ECB_DESCRI"+GetWhere("=",oXls:COL_C)+" AND "+;
                  "ECB_MONTO" +GetWhere("=",oXls:COL_D)+" AND "+;
                  "ECB_SALDO" +GetWhere("=",oXls:COL_E)

      IF !ISSQLFIND(cTable,cWhere)

        oTable:AppendBlank()
        oTable:Replace("ECB_ITEM"  ,STRZERO(nItem,4))
        oTable:Replace("ECB_CODSUC",cCodSuc   )
        oTable:Replace("ECB_FECHA" ,oXls:COL_A)
        oTable:Replace("ECB_REFERE",oXls:COL_B)
        oTable:Replace("ECB_DESCRI",oXls:COL_C)
        oTable:Replace("ECB_MONTO" ,oXls:COL_D)
        oTable:Replace("ECB_SALDO" ,oXls:COL_E)
        oTable:Replace("ECB_CODBCO",oParXls:cCodBco)
        oTable:Replace("ECB_CTABCO",oParXls:cCodCta)
        oTable:Replace("ECB_NUMREG",oParXls:cNumero)
        oTable:Commit("")

        dDesde:=MIN(dDesde,oXls:COL_A)
        dHasta:=MAX(dHasta,oXls:COL_A)

      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

   cWhere:=NIL

   EJECUTAR("BREDOBCOLEE",cWhere,cCodSuc,oParXls:cCodBco,oParXls:cCodCta,oParXls:cNumero,oDp:nIndefinida,dDesde,dHasta,cTitle)

RETURN .T.
