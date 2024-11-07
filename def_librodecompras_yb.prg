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
   LOCAL oTable,oXls,lBrowse
   LOCAL nLinIni,nContar,I,U,nCxP
   LOCAL dDesde:=oParXls:dDesde
   LOCAL dHasta:=oParXls:dHasta

   DEFAULT cCodigo:="LIBRODECOMPRAS_YB"

   oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
   cFileXls:=ALLTRIM(oTable:IXL_FILE  )
   cTable  :=ALLTRIM(oTable:IXL_TABLA )
   nLinIni :=MAX(oTable:IXL_LININI,1)
   lBrowse :=oTable:IXL_LININI
   oTable:End(.T.)

   cWhere:="LBC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
           "LBC_FCHDEC"+GetWhere("=",dHasta)+" AND "+;
           "LBC_ORIGEN"+GetWhere("=","XLS")

   SQLDELETE(cTable,cWhere)

   IF(ValType(oMemo)="O",oMemo:Append("Leyendo "+cFileXls+CRLF),NIL)


   oXls   :=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

   IF oParXls:lBrowse 
     ViewArray(oXls:aData)
   ENDIF

   ADEPURA(oXls:aData,{|a,n| Empty(a[1])})

   IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Fecha "+DTOC(dHasta)+CRLF),NIL)

   oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
   oTable:lAuditar:=.F.
   oTable:SetForeignkeyOff()

   oXls:Gotop()

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

      cCodPro:=STRTRAN(ALLTRIM(oXls:COL_H),"-","")
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
         cCodPro:=EJECUTAR("DPPROVEEDORCREA",cCodPro,oXls:COL_G,cCodPro)
      ENDIF

      cItem:=SQLINCREMENTAL(cTable,"LBC_NUMPAR",cWhere,NIL,NIL,.T.,5)

      oXls:COL_B:=CTOO(oXls:COL_B,"D")

      IF MONTH(dHasta)<>MONTH(oXls:COL_B) .AND. YEAR(dHasta)<>MONTH(oXls:COL_B)
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cTipDoc+"-"+oXls:COL_F+" Periodo Diferente "+CTOO(oXls:COL_B,"C")+CRLF),NIL)
         oXls:DbSkip()   
         LOOP
      ENDIF

      oXls:COL_C:=CTOO(oXls:COL_C,"C") // Viene numérico
      oXls:COL_D:=CTOO(oXls:COL_D,"C") // Viene numérico

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cTipDoc+"-"+oXls:COL_C+CRLF),NIL)
      oTable:AppendBlank()
      oTable:Replace("LBC_NUMFAC",oXls:COL_C)
      oTable:Replace("LBC_NUMFIS",oXls:COL_D)
      oTable:Replace("LBC_ORIGEN","XLS")
      oTable:Replace("LBC_CODIGO",cCodPro)
      oTable:Replace("LBC_RIF",cCodPro)
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


//      oTable:Replace("LBC_MTOBAS",oXls:COL_S)
//      oTable:Replace("LBC_MTOIVA",oXls:COL_T)
      oTable:Replace("LBC_MTONET",oXls:COL_M)
      oTable:Replace("LBC_PORIVA",oXls:COL_S) // *100)
      oTable:Replace("LBC_CODMOD",oDp:cCtaMod)
      oTable:Replace("LBC_MTOEXE",0)
      oTable:SetDefault()
      oTable:Commit("")

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)

   IF(ValType(oMemo )="O",oMemo:Append("Asignando Cuentas Contables"+CRLF),NIL)
   EJECUTAR("DPLIBCOMPRASDETCTAS") // Asigna las cuentas contables desde importar datos desde excel

   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

RETURN .T.
