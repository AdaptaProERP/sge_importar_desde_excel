/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código FERRETERO
// Fecha  19/07/2025
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cWhere
    LOCAL oTable,oXls,oMovInv,oDb:=OpenOdbc(oDp:cDsnData)
    LOCAL nLinIni,nLinFin,nContar,I,U
    LOCAL aRef:={}
    LOCAL cCodGru:="",cCodAlm:="",oGrupo,oAlmacen,nCosto:=0,nPrecio:=0,cUndMed:="UND",nExiste:=0,cLote:="",dFchVen:=CTOD(""),nCostoD:=0,nCxUnd:=1
    LOCAL cCodSuc:=oDp:cSucursal,nCantid:=0,lCheck:=.T.

    DEFAULT cCodigo:="FERRETERO"


    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_LINFIN,IXL_MEMO FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    nLinFin :=MAX(oTable:IXL_LINFIN,0)

//    nLinFin:=10 

    aRef    :=EJECUTAR("IXLLOAD",NIL,NIL,NIL,NIL,oTable:IXL_MEMO)
    AEVAL(aRef,{|a,n| aRef[n,1]:=STRTRAN(a[1],"@","")})
    oTable:End(.T.)

    cWhere :="INV_APLORG"+GetWhere("=","XLS")
    nCantid:=COUNT(cTable,cWhere)

    IF nCantid>0 .AND. MsgYesNo("Desea Remover "+LSTR(nCantid)+" Registros de la tabla "+GetFromVar("{oDp:"+cTable+"}"))
      SQLDELETE(cTable,cWhere)
      lCheck:=.F.
    ENDIF

    lCheck:=IF(nCantid=0,.F.,lCheck)

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni,nLinFin) // nLinIni)
 
    // IF(oXls:lBrowse,oXls:Browse(),NIL)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    // oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
    oTable:=INSERTINTO(cTable,oDb,50)

    IF "TTABLE"$oTable:ClassName()

      IF Empty(oTable:aDefault)
        EJECUTAR("DPTABLESETDEF",oTable) // Facilita Asinar valores por defecto  en :AppendBlank()
      ENDIF

      oTable:lAuditar:=.F.
      oTable:SetForeignkeyOff()

    ENDIF

    oGrupo  :=OpenTable("SELECT * FROM DPGRU"    ,.F.) // Agregar Grupos
    oAlmacen:=OpenTable("SELECT * FROM DPALMACEN",.F.) // Agregar Almacen
    // oMovInv :=OpenTable("SELECT * FROM DPMOVINV" ,.F.) // Agregar Existencia

    oMovInv:=INSERTINTO("DPMOVINV",oDb,100)

    oXls:Gotop()

    WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      // Asigna los valores desde oXls->oRef oReg:RIF = oXls:A
      AEVAL(aRef,{|a,n| aRef[n,2]:=MacroEje("oXls:COL_"+a[3]),;
                        oRef:Set(aRef[n,1],aRef[n,2])})

      cCodigo:=STRTRAN(ALLTRIM(oXls:COL_B),"-","")
      cCodigo:=STRTRAN(cCodigo,";","")
      cCodigo:=STRTRAN(cCodigo,"/","")

      IF Empty(cCodigo)
         oXls:DbSkip()
         LOOP
      ENDIF

      cCosto    :=120  // Costo es necesario
      nCostoD   :=1    // Costo en $

      oXls:COL_A:=CHR(63+CTOO(oXls:COL_A,"N")) // Aseguramos que sea Caracter

      cWhere    :="INV_CODIGO"+GetWhere("=",cCodigo)

      IF !lCheck .OR. !ISSQLFIND(cTable,cWhere)

        IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodigo+CRLF),NIL)

        cCodGru   :=EJECUTAR("DPGRU_CREA",NIL,oXls:COL_J,oGrupo) 

        oTable:AppendBlank()
        oTable:SetDefault() // Asigna valores por defecto

        // AEVAL(aRef,{|a,n| oTable:Replace(a[1],a[2])})

        oTable:Replace("INV_CODIGO",cCodigo)

        // AEVAL(aRef,{|a,n| oTable:Replace(a[1],a[2])})
        oTable:Replace("INV_CODIGO",cCodigo)
        oTable:Replace("INV_DESCRI",oXls:COL_C)
        // oTable:Replace("INV_CATABC",oXls:COL_A)
        oTable:Replace("INV_IVA"   ,"GN")
        oTable:Replace("INV_UTILIZ","V") // Venta
        oTable:Replace("INV_ESTADO","A") // Activo
        oTable:Replace("INV_METCOS","P") // promedio
        oTable:Replace("INV_EDITAR","N") // no editar en facturación
        oTable:Replace("INV_EXISTE",0      ) // Esto debe realizarlo el disparador 
        oTable:Replace("INV_COSPRO",nCosto )
        oTable:Replace("INV_COSADQ",nCosto )
        oTable:Replace("INV_APLORG","XLS"  )  
        oTable:Replace("INV_COSFOB",nCostoD)
        oTable:Replace("INV_COSMER",nCosto )
        oTable:Replace("INV_GRUPO",cCodGru )
        oTable:Replace("INV_CODCAR",oDp:cCodCar)
        oTable:Commit()

      ELSE
        // Actualiza el registro en caso de ser necesario
        // SQLUPDATE(cTable,{"FIELD1","FIELD2"},{oXls:COL_B,oXls:COL_C},cWhere)
      ENDIF

      oXls:DbSkip()

   ENDDO

   /*
   // Existencias
   */
   IF(ValType(oMemo)="O",oMemo:Append("Registrando Existencias"+CRLF),NIL)

   oXls:Gotop()
   WHILE !oXls:Eof()

      cCodigo:=STRTRAN(ALLTRIM(oXls:COL_B),"-","")
      cCodigo:=STRTRAN(cCodigo,";","")
      cCodigo:=STRTRAN(cCodigo,"/","")

      IF Empty(cCodigo)
         oXls:DbSkip()
         LOOP
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodigo+CRLF),NIL)

      cCodAlm :=EJECUTAR("DPALMACENCREA",oXls:COL_F,oXls:COL_N,oDp:cSucursal,oAlmacen)
      nExiste :=CTOO(oXls:COL_H,"N") // Existe

      IF nExiste<>0 

         EJECUTAR("DPINVEXIINI",cCodigo,nExiste,nCosto,cUndMed,nCxUnd,nCostoD,cLote,dFchVen,cCodSuc,cCodAlm,oMovInv,"XLS",.F.,.F.)
   
      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oGrupo:End(.T.)
   oAlmacen:End(.T.)
   oMovInv:End(.T.)

   oTable:End(.T.)
   oXls:End()

RETURN .T.
// EOF
