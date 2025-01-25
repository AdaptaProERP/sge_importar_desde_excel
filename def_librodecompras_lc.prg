/*
// Importación de datos desde excel para MaryPlant
// Tabla  DPLIBCOMPRASDET         
// Código LIBRODEVENTAS_MP
// Fecha  23/03/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
   LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere,cTipDoc,cItem,cNumPar,cCodPro,cItem,cTipIva
   LOCAL oTable,oXls
   LOCAL nLinIni,nContar,I,U,nCxP
   LOCAL dDesde:=oParXls:dDesde
   LOCAL dHasta:=oParXls:dHasta

   DEFAULT cCodigo:="LIBRODECOMPRAS_LC"

   oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
   cFileXls:=ALLTRIM(oTable:IXL_FILE  )
   cTable  :=ALLTRIM(oTable:IXL_TABLA )
   nLinIni :=MAX(oTable:IXL_LININI,1)
   oTable:End(.T.)

   cWhere:="LBC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
           "LBC_FCHDEC"+GetWhere("=",dHasta)+" AND "+;
           "LBC_ORIGEN"+GetWhere("=","XLS")

   SQLDELETE(cTable,cWhere)

   IF(ValType(oMemo)="O",oMemo:Append("Leyendo "+cFileXls+CRLF),NIL)

   oXls   :=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

   ADEPURA(oXls:aData,{|a,n| Empty(a[1])})

   IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Fecha "+DTOC(dHasta)+CRLF),NIL)

   oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
   oTable:lAuditar:=.F.
   oTable:SetForeignkeyOff()

   oXls:Gotop()
   nContar:=0

   WHILE !oXls:Eof()

      cTipDoc:="FAC"
      nCxP   :=1
      cTipIva:="GN"

      IF oXls:COL_O<0
         cTipDoc:="CRE"
         nCxP   :=1
      ENDIF

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cCodPro:=STRTRAN(ALLTRIM(oXls:COL_C),"-","")
      cCodPro:=STRTRAN(cCodPro,"-","")
     
      IF Empty(cCodPro) 
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"-> Línea Vacia"+CRLF),NIL)
         oXls:DbSkip()   
         LOOP
      ENDIF

      IF Empty(oXls:COL_B) 
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"-> Fecha Vacia COL_B"+CRLF),NIL)
         oXls:DbSkip()   
         LOOP
      ENDIF

     
     

      IF !ISSQLFIND("DPPROVEEDOR","PRO_RIF"+GetWhere("=",cCodPro))
         cCodPro:=EJECUTAR("DPPROVEEDORCREA",cCodPro,oXls:COL_D,cCodPro)
      ENDIF

      cItem:=SQLINCREMENTAL(cTable,"LBC_NUMPAR",cWhere,NIL,NIL,.T.,5)
      // nContar:=nContar+1
      // cItem:=STRZERO( nContar:=nContar+1

      oXls:COL_B:=CTOO(oXls:COL_B,"D")

      IF MONTH(dHasta)<>MONTH(oXls:COL_B) .AND. YEAR(dHasta)<>MONTH(oXls:COL_B)
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cTipDoc+"-"+oXls:COL_F+" Periodo Diferente "+CTOO(oXls:COL_B,"C")+CRLF),NIL)
        // oXls:DbSkip()   
        // LOOP
      ENDIF

      oXls:COL_I:=CTOO(oXls:COL_I,"C") // Viene numérico
      oXls:COL_J:=CTOO(oXls:COL_J,"C") // Viene numérico

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cTipDoc+"-"+oXls:COL_I+CRLF),NIL)
      oTable:AppendBlank()
      oTable:Replace("LBC_NUMFAC",oXls:COL_F)
      oTable:Replace("LBC_NUMFIS",oXls:COL_G)
      oTable:Replace("LBC_ORIGEN","XLS")
      oTable:Replace("LBC_CODIGO",cCodPro)
      oTable:Replace("LBC_RIF"   ,cCodPro)
      oTable:Replace("LBC_FCHDEC",dHasta)
      oTable:Replace("LBC_FECHA" ,oXls:COL_B)
      oTable:Replace("LBC_TIPDOC",cTipDoc)
      oTable:Replace("LBC_CODSUC",oDp:cSucursal)
      oTable:Replace("LBC_USUARI",oDp:cUsuario)
      oTable:Replace("LBC_NUMPAR",cItem)
      oTable:Replace("LBC_ACTIVO",.T.)
      oTable:Replace("LBC_CXP"   ,nCxP)
      oTable:Replace("LBC_ITEM"  ,"00001")
      oTable:Replace("LBC_TIPIVA",cTipIva)
      oTable:Replace("LBC_COMORG","N"    ) // Nacional
      oTable:Replace("LBC_CODMON",oDp:cMonedaExt ) 

      oTable:Replace("LBC_BASGN" ,oXls:COL_L)
      oTable:Replace("LBC_BASEX" ,oXls:COL_K)

      oTable:Replace("LBC_MTOIVA",oXls:COL_N)
      oTable:Replace("LBC_MTONET",oXls:COL_O)
//    oTable:Replace("LBC_PORIVA",oXls:COL_R*100)
      oTable:Replace("LBC_CODMOD",oDp:cCtaMod)

      oTable:Replace("LBC_NUMRTI",RIGTH(oXls:COL_V,8))
      oTable:Replace("LBC_MTORTI",oXls:COL_W)

      oTable:Replace("LBC_MTOEXE",0)
      oTable:SetDefault()
      oTable:Commit("")

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)

   IF(ValType(oMemo )="O",oMemo:Append("Asignando Cuentas Contables"+CRLF),NIL)
   EJECUTAR("DPLIBCOMPRASDETCTAS") // Asigna las cuentas contables desde importar datos desde excel

   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

// EJECUTAR("BRLIBCOMEDIT",NIL,cCodSuc,NIL,dDesde,dHasta,NIL,dFchDec,lView,oLIBCOMFCH:cCodCaj,NIL,NIL,NIL,NIL,oLIBCOMFCH:cCenCos,aTipDoc,lCondom,lCtaEgr,oLIBCOMFCH:lVenta)


   oTable:End(.T.)
   oXls:End()

RETURN .T.
