/*
// Importaci�n de datos desde excel
// Tabla  <TABLA>
// C�digo TRABAJADOR
// Fecha  23/11/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere
    LOCAL oTable,oXls
    LOCAL nLinIni,nContar,I,U
    LOCAL aRef:={}

    DEFAULT cCodigo:="TRABAJADOR"

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_MEMO FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    aRef    :=EJECUTAR("IXLLOAD",NIL,NIL,NIL,NIL,oTable:IXL_MEMO)
    AEVAL(aRef,{|a,n| aRef[n,1]:=STRTRAN(a[1],"@","")})
    oTable:End(.T.)


    IF COUNT(cTable)>0 .AND. MsgYesNo("Desea Remover todos los Registros de la tabla "+cTable)
      SQLDELETE(cTable)
    ENDIF

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

    IF(oImpXls:lBrowse,oXls:Browse(),NIL)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    oTable:=OpenTable("SELECT * FROM "+cTable, .F. )

    IF Empty(oTable:aDefault)
       EJECUTAR("DPTABLESETDEF",oTable) // Facilita Asinar valores por defecto  en :AppendBlank()
    ENDIF

    oTable:lAuditar:=.F.
    oTable:SetForeignkeyOff()

    oXls:Gotop()

    WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      // Asigna los valores desde oXls->oRef oReg:RIF = oXls:A
      AEVAL(aRef,{|a,n| aRef[n,2]:=MacroEje("oXls:COL_"+a[3]),;
                        oRef:Set(aRef[n,1],aRef[n,2])})

      cCodigo:=CTOO(oXls:COL_B,"C")
      cCodigo:=STRTRAN(ALLTRIM(oXls:COL_B),"-","")
      cCodigo:=STRTRAN(cCodigo,";","")
      cCodigo:=STRTRAN(cCodigo,"/","")

      IF Empty(cCodigo)
         oXls:DbSkip()
         LOOP
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodigo+CRLF),NIL)

      cWhere    :="CODIGO"+GetWhere("=",cCodigo)

      IF !ISSQLFIND(cTable,cWhere)
        oTable:AppendBlank()
        AEVAL(aRef,{|a,n| oTable:Replace(a[1],a[2])})
        oTable:Replace("CODIGO",cCodigo)
        oTable:Replace("TIPO_NOM","Q")
        oTable:Replace("CONDICION","A")
        oTable:Replace("TRA_ACTIVO",.T.)
        oTable:Commit("")
      ELSE
        // Actualiza el registro en caso de ser necesario
        SQLUPDATE(cTable,{"CODIGO"},{cCodigo},cWhere)
      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importaci�n Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

   EJECUTAR("NMTRABAJADORFIX")
   EJECUTAR("NMTRABAJADOR")

RETURN .T.
