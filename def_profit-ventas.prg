/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código LIBRODEVENTAS_MP
// Fecha  23/03/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
   LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere,cTipDoc,cItem,cNumPar,cCodCli,cItem,cTipIva
   LOCAL oTable,oXls
   LOCAL nLinIni,nContar,I,U,nCxC
   LOCAL dDesde:=oParXls:dDesde
   LOCAL dHasta:=oParXls:dHasta

   DEFAULT cCodigo:="LIBRODEVENTAS_MP"

   oTable:=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
   cFileXls:=ALLTRIM(oTable:IXL_FILE  )
   cTable  :=ALLTRIM(oTable:IXL_TABLA )
   nLinIni :=MAX(oTable:IXL_LININI,1)
   oTable:End(.T.)

   cWhere:="LBC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
           "LBC_FCHDEC"+GetWhere("=",dHasta)+" AND "+;
           "LBC_ORIGEN"+GetWhere("=","XLS")

   SQLDELETE(cTable,cWhere)

//   IF oDp:oXls=NIL .OR. Empty(oDp:oXls:aData)
      oXls   :=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)
//      oDp:oXls:=oXls
//   ELSE
      oXls:=oDp:oXls
//   ENDIF

   ADEPURA(oXls:aData,{|a,n| Empty(a[1])})
//   ViewArray(oXls:aData)

   IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

   oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
   oTable:lAuditar:=.F.
   oTable:SetForeignkeyOff()

   oXls:Gotop()

   WHILE !oXls:Eof()

      cTipDoc:="FAV"
      cTipIva:="GN"

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cCodCli:=STRTRAN(ALLTRIM(oXls:COL_D),"-","")
      cCodCli:=STRTRAN(cCodCli,"-","")
     
      IF Empty(cCodCli)
         oXls:DbSkip()   
         LOOP
      ENDIF

      IF !ISSQLFIND("DPCLIENTES","CLI_RIF"+GetWhere("=",cCodCli))
         cCodCli:=EJECUTAR("DPCLIENTECREA",cCodCli,oXls:COL_C,cCodCli)
      ENDIF

      cItem:=SQLINCREMENTAL(cTable,"LBC_NUMPAR",cWhere,NIL,NIL,.T.,5)

      IF MONTH(dHasta)<>MONTH(oXls:COL_B) .AND. YEAR(dHasta)<>MONTH(oXls:COL_B)
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cTipDoc+"-"+oXls:COL_F+" Periodo Diferente "+DTOC(oXls:COL_B)+CRLF),NIL)
         oXls:DbSkip()   
         LOOP
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cTipDoc+"-"+oXls:COL_F+CRLF),NIL)
      oTable:AppendBlank()
      oTable:Replace("LBC_NUMFAC",oXls:COL_F)
      oTable:Replace("LBC_NUMFIS",oXls:COL_G)
      oTable:Replace("LBC_ORIGEN","XLS")
      oTable:Replace("LBC_CODIGO",cCodCli)
      oTable:Replace("LBC_RIF",cCodCli)
      oTable:Replace("LBC_FCHDEC",dHasta)
      oTable:Replace("LBC_FECHA" ,oXls:COL_B)
      oTable:Replace("LBC_TIPDOC",cTipDoc)
      oTable:Replace("LBC_CODSUC",oDp:cSucursal)
      oTable:Replace("LBC_USUARI",oDp:cUsuario)
      oTable:Replace("LBC_NUMPAR",cItem)
      oTable:Replace("LBC_ACTIVO",.T.)
      oTable:Replace("LBC_CXC"   ,nCxC)
      oTable:Replace("LBC_ITEM"  ,"0001")
      oTable:Replace("LBC_TIPIVA",cTipIva)

      oTable:Replace("LBC_MTOBAS",oXls:COL_N)
      oTable:Replace("LBC_MTOIVA",oXls:COL_P)
      oTable:Replace("LBC_MTONET",oXls:COL_L)
      oTable:Replace("LBC_PORIVA",oXls:COL_O*100)
      oTable:Replace("LBC_MTOEXE",0)
      oTable:Commit("")

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
  // oXls:End()

RETURN .T.
