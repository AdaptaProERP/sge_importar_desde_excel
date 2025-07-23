/*
// Importación de datos desde excel para MaryPlant
// Tabla  DPLIBVENTASDET         
// Código LIBRODEVENTAS_PP_VENTAS
// Fecha  23/05/2025
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
   LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere,cTipDoc,cItem,cNumPar,cCodCli,nItem:=0,cTipIva,cTitle
   LOCAL oTable,oXls
   LOCAL nLinIni,nLinFin,nContar:=0,I,U,nCxC,nAt
   LOCAL dDesde :=oParXls:dDesde
   LOCAL dHasta :=oParXls:dHasta
   LOCAL dFchDec:=dHasta,cFecha
   LOCAL nT1
   LOCAL cCtaIgtf,cCtaCaj:="",cCtaAnt:="",cCtaEgr:=oDp:cCtaIndef
   LOCAL cCtaVta:=""
   LOCAL nNumFile:=0
   LOCAL cNumPar :="",cFecha,cRifLib
   LOCAL cNumDoc,cFacAfe

   DEFAULT cCodigo:="LIBRODEVENTAS_PP_VENTAS"

   IF COUNT("DPCTA")<=1
      MsgMemo("Debe Importar Plan de Cuentas")
      EJECUTAR("DPCTAIMPORT")
      RETURN .T.
   ENDIF

   oDp:cCtaVtaNac:=SQLGET("VIEW_DPCODINTEGRA","CIN_CODCTA,CIN_ABREVI","CIN_CODIGO='VTANAC'"    )

   oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_LINFIN FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
   cFileXls:=ALLTRIM(oTable:IXL_FILE  )
   cTable  :=ALLTRIM(oTable:IXL_TABLA )
   cTable  :="dplibventasdet"
   nLinIni :=MAX(oTable:IXL_LININI,1)
   nLinFin :=MAX(oTable:IXL_LINFIN,0)

   oTable:End(.T.)

   nNumFile:=EJECUTAR("DPFILEEMPGETNUM",cFileXls)
   cWhere  :="LBC_NUMFIL"+GetWhere("=",nNumFile)
   nT1     :=SECONDS()

   IF nNumFile>0

     IF !EJECUTAR("XLSDELETEDOCPRO",nNumFile,NIL,.T.,.T.)
        RETURN .F.
     ENDIF

   ENDIF

   IF(ValType(oMemo)="O",oMemo:Append("Leyendo "+cFileXls+CRLF),NIL)

   oXls   :=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,2,2,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,.T.)  
   cRifLib:=oXls:A3
   cFecha :=oXls:E2
   nAt    :=AT("Desde",cFecha)
   cFecha :=SUBS(cFecha,nAt+6,LEN(cFecha))
   cFecha :=STRTRAN(cFecha,"Hasta","")
   cFecha :=ALLTRIM(STRTRAN(cFecha,")",""))
   cFecha :=ALLTRIM(STRTRAN(cFecha,"-","/"))

   dDesde :=CTOD(LEFT(cFecha,10))
   dHasta :=CTOD(RIGHT(cFecha,10))
   dFchDec:=dHasta
   oXls:End()

   IF(ValType(oMemo)="O",oMemo:Append("Periodo "+DTOC(dDesde)+"-"+DTOC(dHasta)+CRLF),NIL)

   oXls   :=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni,nLinFin)

   AEVAL(oXls:aData,{|a,n| oXls:aData[n,2]:=CTOO(a[2],"C"),;
                           oXls:aData[n,1]:=CTOO(a[1],"C")})

//   ADEPURA(oXls:aData,{|a,n| Empty(a[2]) .OR. "FECHA"$a[2] .OR. "t"$a[1]}) // Sin fecha no es necesario

   IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Periodo "+DTOC(dDesde)+"-"+DTOC(dHasta)+CRLF),NIL)
   IF(ValType(oMemo )="O",DPFOCUS(oMemo),NIL)
   oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
   oTable:lAuditar:=.F.
   oTable:SetForeignkeyOff()
   nContar:=0

//   dDesde:=CTOD("")
//   dHasta:=CTOO("")

   oXls:Gotop()

   WHILE !oXls:Eof()

      cTipDoc:="FAV"
      nCxC   :=1
      cTipIva:="GN"
      cCodCta:=oDp:cCtaVtaNac
      cNumDoc:=CTOO(oXls:COL_F,"C")

      IF !Empty(oXls:COL_E)
        cTipDoc:="RTI" // retencion de IVA
        cNumDoc:=ALLTRIM(CTOO(oXls:COL_E,"C"))
        cNumDoc:=RIGHT(cNumDoc,8)
      ENDIF

      IF !Empty(oXls:COL_H)
        cTipDoc:="DEB" // retencion de IVA
        cNumDoc:=CTOO(oXls:COL_H,"C")
      ENDIF



      cFacAfe:=CTOO(oXls:COL_K,"C")


/*
      oXls:COL_M:=CTOO(oXls:COL_M,"N") // Monto Neto

      IF oXls:COL_M<0
         cTipDoc:="CRE"
         nCxC   :=1
      ENDIF
*/

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cCodCli:=STRTRAN(ALLTRIM(oXls:COL_C),"-","")
      cCodCli:=STRTRAN(cCodCli,"-","")
      cCodCli:=ALLTRIM(cCodCli)
      
      IF Empty(cCodCli) 
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"-> Línea Vacia"+CRLF),NIL)
         oXls:DbSkip()   
         LOOP
      ENDIF

      IF LEN(cCodCli)>10
         IF(ValType(oSay  )="O",oSay:SetText("RIF:"+cCodCli+" tiene "+LSTR(LEN(cCodCli))+" Dígitos"),nil)
      ENDIF

      cCodCli:=LEFT(cCodCli,10)
      cFecha :=CTOO(oXls:COL_B,"C")

      oXls:COL_B:=CTOO(oXls:COL_B,"D")

      IF Empty(oXls:COL_B) 
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"-> Fecha "+cFecha+" Vacia COL_B"+CRLF),NIL)
         oXls:DbSkip()   
         LOOP
      ENDIF

      IF !ISSQLFIND("DPCLIENTES","CLI_RIF"+GetWhere("=",cCodCli))
         cCodCli:=EJECUTAR("DPCLIENTECREA",cCodCli,oXls:COL_D,cCodCli)
      ENDIF

      nContar++
      cItem     :=STRZERO(1,5)  // nContar,5) // SQLINCREMENTAL(cTable,"LBC_NUMPAR",cWhere,NIL,NIL,.T.,5)
      cNumPar   :=SQLINCREMENTAL(cTable,"LBC_NUMPAR",NIL,NIL,NIL,.T.,5)
      

      // oXls:COL_B:=CTOO(oXls:COL_B,"D")

      // #CONTROL
      oXls:COL_G:=CTOO(oXls:COL_G,"C") // NUMERO DE CONTROL
      oXls:COL_F:=CTOO(oXls:COL_F,"C") // NUMERO DE FACTURA

      IF oDp:cTipCon="F"
        // dFchDec:=EJECUTAR("GETFCHTRIMESTRE",oXls:COL_B)
      ENDIF

// ? oXls:COL_F,"<-F",oXls:COL_G,"oXls:COL_G"
//      IF !ISSQLFIND("DPCTA","CTA_CODIGO"+GetWhere("=",oXls:COL_N))
//        EJECUTAR("DPCTACREA",oXls:COL_N)
//      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cTipDoc+"-"+ALLTRIM(oXls:COL_E)+" "+cFecha+CRLF),NIL)
      oTable:AppendBlank()
      oTable:Replace("LBC_NUMFAC",cNumDoc   )
      oTable:Replace("LBC_NUMFIS",oXls:COL_G)
      oTable:Replace("LBC_ORIGEN","XLS"     )
      oTable:Replace("LBC_CODIGO",cCodCli)
      oTable:Replace("LBC_RIF"   ,cCodCli)
      oTable:Replace("LBC_FCHDEC",dFchDec     )
      oTable:Replace("LBC_FACAFE",cFacAfe     )

      oTable:Replace("LBC_FECHA" ,oXls:COL_B  )
      oTable:Replace("LBC_TIPDOC",cTipDoc     )
      oTable:Replace("LBC_CODSUC",oParXls:cCodSuc)
      oTable:Replace("LBC_USUARI",oDp:cUsuario)
      oTable:Replace("LBC_NUMPAR",cNumPar     )
      oTable:Replace("LBC_ACTIVO",.T.         )
      oTable:Replace("LBC_CXC"   ,nCxC        )
      oTable:Replace("LBC_ITEM"  ,cItem       )
      oTable:Replace("LBC_TIPIVA",cTipIva     )
      oTable:Replace("LBC_COMORG","N"         ) // Nacional
      oTable:Replace("LBC_CODMON",oDp:cMoneda )
 
      oTable:Replace("LBC_BASGN" ,oXls:COL_R  )
      oTable:Replace("LBC_MTOIVA",oXls:COL_K  )
      oTable:Replace("LBC_MTONET",oXls:COL_L  )
      oTable:Replace("LBC_CODMOD",oDp:cCtaMod )
      oTable:Replace("LBC_CTAEGR",cCtaEgr     )
      oTable:Replace("LBC_CODCTA",cCodCta     )

      oTable:Replace("LBC_NUMFIL",nNumFile    )

      oTable:SetDefault()
      oTable:Commit("")

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)

   IF(ValType(oMemo )="O",oMemo:Append("Asignando Cuentas Contables"+CRLF),NIL)
  //  EJECUTAR("DPLIBVENTASDETCTAS") // Asigna las cuentas contables desde importar datos desde excel

   IF(ValType(oMemo )="O",oMemo:Append("Importación "+LSTR(nContar)+" Registros realizada en "+LSTR(SECONDS()-nT1)+" Segundos" +CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

   IF ValType(oMemo )="O"
     nT1     :=ABS(nT1-SECONDS())
     SQLUPDATE("dpfileemp",{"FIL_TEXTO","FIL_TABLE","FIL_REV"},{oMemo:GetText(),"DPLIBVENTASDET",LSTR(nT1)},"FIL_NUMERO"+GetWhere("=",nNumFile))
   ENDIF

   cTitle:="Libro de VENTAS "+DTOC(dDesde)+" "+DTOC(dHasta)

   IF !Empty(dDesde)
     EJECUTAR("BRLIBCOMEDIT",cWhere,NIL,oDp:nIndefinida,dDesde,dHasta,cTitle,CTOD(""),.F.,NIL,NIL,NIL,NIL,NIL,NIL,NIL,.F.,NIL,.T.)
   ENDIF

RETURN .T.
// EOF
