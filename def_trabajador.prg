/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código TRABAJADOR
// Fecha  23/11/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere
    LOCAL oTable,oXls
    LOCAL nLinIni,nLinFin,nContar,I,U,nNumFile
    LOCAL aRef:={}

    DEFAULT cCodigo:="TRABAJADOR"

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_LINFIN,IXL_MEMO FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)

? oDp:cSql

    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    nLinFin :=oTable:IXL_LINFIN
    aRef    :=EJECUTAR("IXLLOAD",NIL,NIL,NIL,NIL,oTable:IXL_MEMO)

    AEVAL(aRef,{|a,n| aRef[n,1]:=STRTRAN(a[1],"@","")})
    oTable:End(.T.)

    nNumFile:=EJECUTAR("DPFILEEMPGETNUM",cFileXls,cCodigo)

    cWhere  :="TRA_NUMFIL"+GetWhere("=",nNumFile)
    nCantid :=COUNT(cTable,cWhere)

    IF nCantid>0 .AND. MsgYesNo("Desea Remover "+LSTR(nCantid)+" Registros de la tabla "+cTable)
      SQLDELETE(cTable,cWhere)
    ENDIF

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

oXls:Browse()

    IF(oImpXls:lBrowse,oXls:Browse(),NIL)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

? cTable


return 

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
//      ELSE
        // Actualiza el registro en caso de ser necesario
//        SQLUPDATE(cTable,{"CODIGO"},{cCodigo},cWhere)
      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

   EJECUTAR("NMTRABAJADORFIX")
   EJECUTAR("NMTRABAJADOR")

RETURN .T.

FUNCTION GETGRUPO(cNombre)
  LOCAL cCodigo
  cCodigo:=GETNOMBREX("NMGRUPO",cNombre,"GTR_CODIGO","GTR_DESCRI")
RETURN cCodigo

FUNCTION GETDPTO(cNombre)
  LOCAL cCodigo
  cCodigo:=GETNOMBREX("DPDPTO",cNombre,"DEP_CODIGO","DEP_DESCRI")
RETURN cCodigo

FUNCTION GETUND(cNombre)
   LOCAL cCodigo
   cCodigo:=GETNOMBREX("NMUNDFUNC",cNombre,"CEN_CODIGO","CEN_DESCRI")
RETURN cCodigo

FUNCTION GETCARGO(cNombre)
   LOCAL cCodigo
   cCodigo:=GETNOMBREX("NMCARGOS",cNombre,"CAR_CODIGO","CAR_DESCRI")
RETURN cCodigo

FUNCTION GETBANCO(cNombre)
   LOCAL cCodigo
   cCodigo:=GETNOMBREX("NMBANCOS",cNombre,"BAN_CODIGO","BAN_NOMBRE")
RETURN cCodigo


FUNCTION GETNOMBREX(cTable,cNombre,cClave,cDescri,oTable)
  LOCAL cCodigo

  IF EMPTY(cNombre)
     cNombre:="Por Definir"
  ELSE
     cNombre:=ALLTRIM(CTOO(cNombre,"C"))
  ENDIF

  cCodigo:=SQLGET(cTable,cClave,cDescri+GetWhere("=",cNombre))

  IF Empty(cCodigo)

    cCodigo:=SQLGETMAX(cTable,cClave)
    cCodigo:=STRZERO(VAL(cCodigo)+1,LEN(cCodigo))

    WHILE .T.

       IF Empty(SQLGET(cTable,cClave,cClave+GetWhere("=",cCodigo)))
          EXIT
       ENDIF

       cCodigo:=STRZERO(VAL(cCodigo)+1,LEN(cCodigo))

       // ? cCodigo,cTable,cClave,"cCodigo, ya Existe"

    ENDDO

    oTable:=OpenTable("SELECT * FROM "+cTable,.F.)
    oTable:lAuditar:=.F.
    oTable:AppendBlank()
    oTable:Replace(cClave ,cCodigo)
    oTable:Replace(cDescri,cNombre)
    oTable:Commit()
    oTable:End()

  ENDIF

RETURN cCodigo

// EOF



