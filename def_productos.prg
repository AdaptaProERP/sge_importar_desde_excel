/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código PRODUCTOS
// Fecha  19/07/2025
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cWhere
    LOCAL oTable,oXls
    LOCAL nLinIni,nLinFin,nContar,I,U
    LOCAL aRef:={},cCodGru:="",oGrupo

    DEFAULT cCodigo:="FERRETERO"

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_LINFIN,IXL_MEMO FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    nLinFin :=MAX(oTable:IXL_LINFIN,0)

    aRef    :=EJECUTAR("IXLLOAD",NIL,NIL,NIL,NIL,oTable:IXL_MEMO)
    AEVAL(aRef,{|a,n| aRef[n,1]:=STRTRAN(a[1],"@","")})
    oTable:End(.T.)

    IF COUNT(cTable)>0 .AND. MsgYesNo("Desea Remover todos los Registros de la tabla "+cTable)
      SQLDELETE(cTable)
    ENDIF

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni,nLinFin) // nLinIni)

    IF(oImpXls:lBrowse,oXls:Browse(),NIL)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
    IF Empty(oTable:aDefault)
       EJECUTAR("DPTABLESETDEF",oTable) // Facilita Asinar valores por defecto  en :AppendBlank()
    ENDIF
    oTable:lAuditar:=.F.
    oTable:SetForeignkeyOff()

    oGrupo:=OpenTable("SELECT * FROM DPGRU",.F.) // Agregar Grupos

// oXls:Browse()

    oXls:Gotop()

    WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      // Asigna los valores desde oXls->oRef oReg:RIF = oXls:A
//      AEVAL(aRef,{|a,n| aRef[n,2]:=MacroEje("oXls:COL_"+a[3]),;
//                        oRef:Set(aRef[n,1],aRef[n,2])})

      cCodigo:=STRTRAN(ALLTRIM(oXls:COL_B),"-","")
      cCodigo:=STRTRAN(cCodigo,";","")
      cCodigo:=STRTRAN(cCodigo,"/","")

      IF Empty(cCodigo)
         oXls:DbSkip()
         LOOP
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodigo+CRLF),NIL)

      cWhere    :="INV_CODIGO"+GetWhere("=",cCodigo)

      oXls:COL_A:=CHR(63+CTOO(oXls:COL_A,"N")) // Aseguramos que sea Caracter

//    oXls:COL_A:=IF(oXls:COL_A="1","A",oXls:COL_A)
//    oXls:COL_A:=IF(oXls:COL_A="2","B",oXls:COL_A)
//    oXls:COL_A:=IF(oXls:COL_A="3","C",oXls:COL_A)
//    oXls:COL_A:=IF(oXls:COL_A="4","D",oXls:COL_A)

      oXls:COL_J:=CTOO(oXls:COL_J,"C") // Aseguramos que sea Caracter

      cCodGru   :=GETGRUPO(NIL,oXls:COL_J)

//  ? oXls:COL_J,cCodGru
// EXIT

      IF !ISSQLFIND(cTable,cWhere)

        oTable:AppendBlank()
        oTable:Replace("INV_CODIGO",cCodigo)
        oTable:Replace("INV_DESCRI",oXls:COL_C)
        oTable:Replace("INV_CATABC",oXls:COL_A)
        oTable:Replace("INV_IVA"   ,"GN")

        // AEVAL(aRef,{|a,n| oTable:Replace(a[1],a[2])})
        oTable:Replace("INV_GRUPO",cGrupo)

        oTable:Commit("")
      ELSE
        // Actualiza el registro en caso de ser necesario
        // SQLUPDATE(cTable,{"FIELD1","FIELD2"},{oXls:COL_B,oXls:COL_C},cWhere)
      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oGrupo:End()

   oTable:End(.T.)
   oXls:End()

RETURN .T.

/*
// Obtiene Grupo
*/
FUNCTION GETGRUPO(cGrupo,cDescri)
  LOCAL oTable

  IF Empty(cGrupo) .AND. !Empty(cDescri)
     cGrupo:=SQLGET("DPGRU","GRU_CODIGO","GRU_DESCRI"+GetWhere("=",cDescri))
  ENDIF

  IF Empty(cGrupo) .AND. Empty(cDescri)
     cGrupo:=SQLGET("DPGRU","GRU_CODIGO")
  ENDIF

  IF Empty(cGrupo)
     cGrupo:=SQLINCREMENTAL("DPGRU","GRU_CODIGO",NIL,NIL,NIL,.T.,3)
     cGrupo:=IF(Empty(cGrupo),STRZERO(1,3),cGrupo)
  ENDIF

  cGrupo:=ALLTRIM(cGrupo)

  IF ALLTRIM(SQLGET("DPGRU","GRU_CODIGO","GRU_CODIGO"+GetWhere("=",cGrupo)))=cGrupo
     RETURN cGrupo
  ENDIF

  oGrupo:Append()
  oGrupo:lAuditar:=.F.
  oGrupo:Replace("GRU_CODIGO",cGrupo )
  oGrupo:Replace("GRU_DESCRI",cDescri)
  oGrupo:Replace("GRU_ACTIVO",.T.    )
  oGrupo:Commit("")

RETURN cGrupo
