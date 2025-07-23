/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código FRENOS_VITO
// Fecha  23/03/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere
    LOCAL oTable,oXls,aCopia:={}
    LOCAL nLinIni,nContar,I,U

    DEFAULT cCodigo:="FRENOS_VITO"

    oTable:=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    oTable:End(.T.)

    IF COUNT(cTable)>0 .AND. MsgYesNo("Desea Remover todos los Registros de la tabla "+cTable)
      SQLDELETE(cTable)
    ENDIF

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls   :=EJECUTAR("XLSTORDD",cFileXls)

    ADEPURA(oXls:aData,{|a,n| Empty(a[1]) .OR. "refere"$a[1]})

    AEVAL(oXls:aData,{|a,n|IF(LEN(a)<>12,AADD(aCopia,a),NIL)})

    ADEPURA(oXls:aData,{|a,n| Empty(a[1]) .OR. LEN(a)<>12 .OR. "refere"$a[1]})
//? oXls:Reccount()
//    ViewArray(oXls:aData)

//  ViewArray(aCopia)
// RETURN .T.

   IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

   oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
   oTable:lAuditar:=.F.
   oTable:SetForeignkeyOff()

   oXls:Gotop()

   WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cCodigo:=STRTRAN(ALLTRIM(oXls:COL_A),"-","")
      cCodigo:=STRTRAN(cCodigo,";","")
      cCodigo:=STRTRAN(cCodigo,"/","")

      IF Empty(cCodigo)
         oXls:DbSkip()          LOOP
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodigo+CRLF),NIL)

      cWhere    :="FIELD"+GetWhere("=",cCodigo)

      IF !ISSQLFIND(cTable,cWhere)
        oTable:AppendBlank()
        oTable:Replace("FIELD",cCodigo)
        oTable:Commit("")
      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

RETURN .T.
