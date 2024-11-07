/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código MIPLANDECUENTAS
// Fecha  02/04/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere
    LOCAL oTable,oXls
    LOCAL nLinIni,nContar,I,U

    DEFAULT cCodigo:="MIPLANDECUENTAS"

    oTable:=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    oTable:End(.T.)

    IF COUNT(cTable)>0 .AND. MsgYesNo("Desea Remover todos los Registros de la tabla "+cTable)
      SQLDELETE(cTable)
    ENDIF

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
    oTable:lAuditar:=.F.
    oTable:SetForeignkeyOff()

    oXls:Gotop()

    WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cCodigo:=CTOO(oXls:COL_A,"C")
      cCodigo:=STRTRAN(cCodigo,"-","")
      cCodigo:=STRTRAN(cCodigo,";","")
      cCodigo:=STRTRAN(cCodigo,"/","")

      IF Empty(cCodigo)
         oXls:DbSkip()
         LOOP
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodigo+CRLF),NIL)

      cWhere    :="CTA_CODIGO"+GetWhere("=",cCodigo)

      IF !ISSQLFIND(cTable,cWhere)
        oTable:AppendBlank()
        oTable:Replace("CTA_CODIGO",cCodigo)
        oTable:Replace("CTA_DESCRI",oXls:COL_B)
        oTable:Replace("CTA_ACTIVO",.T.)
        oTable:Commit("")

        // crear código de Integración
        IF !Empty(oTable:CTA_CODINT)
           EJECUTAR("SETCTAINTMOD","DPCODINTEGRA_CTA",oTable:CTA_CODINT,"","CODCTA",oTable:CTA_CODIGO,.T.)
        ENDIF

      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Asignando Cuentas de Movimiento"+CRLF),NIL)
   EJECUTAR("SETDPCTADET")
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

   EJECUTAR("DPCTAUSO")

RETURN .T.
