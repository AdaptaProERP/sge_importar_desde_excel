/*
// Importación de datos desde excel para MaryPlant
// Tabla  DPLIBCOMPRASDET         
// Código LIBRODEVENTAS_MP
// Fecha  23/03/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
   LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere,cTipDoc,cItem,cNumPar,cCodPro,nItem:=0,cTipIva,cTitle
   LOCAL oTable,oXls
   LOCAL nLinIni,nLinFin,nContar:=0,I,U,nCxP
   LOCAL dDesde :=oParXls:dDesde
   LOCAL dHasta :=oParXls:dHasta
   LOCAL dFchDec:=dHasta
   LOCAL nT1
   LOCAL cCtaIgtf,cCtaCaj:="",cCtaAnt:="",cCtaEgr:=oDp:cCtaIndef
   LOCAL nNumFile:=0
   LOCAL cNumPar :="",cFecha

   DEFAULT cCodigo:="COMPRAS COLEGIO CAPACHO"

   IF COUNT("DPCTA")<=1
      MsgMemo("Debe Importar Plan de Cuentas")
      EJECUTAR("DPCTAIMPORT")
      RETURN .T.
   ENDIF

   oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_LINFIN FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
   cFileXls:=ALLTRIM(oTable:IXL_FILE  )
   cTable  :=ALLTRIM(oTable:IXL_TABLA )
   nLinIni :=MAX(oTable:IXL_LININI,1)
   nLinFin :=MAX(oTable:IXL_LINFIN,0)

   oTable:End(.T.)

   nNumFile:=EJECUTAR("DPFILEEMPGETNUM",cFileXls)
   cWhere  :="LBC_NUMFIL"+GetWhere("=",nNumFile)
   nT1     :=SECONDS()

   IF nNumFile>0

     IF !EJECUTAR("XLSDELETEDOCPRO",nNumFile,NIL,NIL,.T.)
        RETURN .F.
     ENDIF

   ENDIF

   IF(ValType(oMemo)="O",oMemo:Append("Leyendo "+cFileXls+CRLF),NIL)

   // LEER LA CUENTAS CONTABLES 
   oXls   :=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,2,4,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,.T.)

   cCtaCaj :=CTOO(oXls:O3,"C")
   cCtaAnt :=CTOO(oXls:P3,"C")
   cCtaIgtf:=CTOO(oXls:L2,"C")

   // ARCHIVO 2023 tiene la indicacion de las cuentas en columnsa diferentes
   IF Empty(cCtaCaj)
     cCtaCaj :=CTOO(oXls:O4,"C")
     cCtaAnt :=CTOO(oXls:P4,"C")
     cCtaIgtf:=CTOO(oXls:L3,"C")
   ENDIF

   IF !Empty(cCtaIgtf)
      EJECUTAR("DPCODINTEGRA_ADD","IGTFPAG","IGTF Pagado",cCtaIgtf)
   ENDIF

   IF !Empty(cCtaAnt)
      EJECUTAR("DPCODINTEGRA_ADD","COMANT","Anticipos Proveedores",cCtaAnt)
   ENDIF

   IF !Empty(cCtaCaj)
      EJECUTAR("DPCODINTEGRA_ADD","CAJEXT","Caja moneda Extranjera",cCtaCaj)
   ENDIF

   oXls:End()

   oXls   :=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni,nLinFin)

   AEVAL(oXls:aData,{|a,n| oXls:aData[n,2]:=CTOO(a[2],"C"),;
                           oXls:aData[n,1]:=CTOO(a[1],"C")})

   ADEPURA(oXls:aData,{|a,n| Empty(a[2]) .OR. "FECHA"$a[2] .OR. "t"$a[1]}) // Sin fecha no es necesario

   IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Periodo "+DTOC(dDesde)+"-"+DTOC(dHasta)+CRLF),NIL)
   IF(ValType(oMemo )="O",DPFOCUS(oMemo),NIL)
   oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
   oTable:lAuditar:=.F.
   oTable:SetForeignkeyOff()
   nContar:=0

   dDesde:=CTOD("")
   dHasta:=CTOO("")

   oXls:Gotop()

   WHILE !oXls:Eof()

      cTipDoc:="FAC"
      nCxP   :=1
      cTipIva:="GN"

      oXls:COL_M:=CTOO(oXls:COL_M,"N") // Monto Neto

      IF oXls:COL_M<0
         cTipDoc:="CRE"
         nCxP   :=1
      ENDIF

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cCodPro:=STRTRAN(ALLTRIM(oXls:COL_F),"-","")
      cCodPro:=STRTRAN(cCodPro,"-","")
      cCodPro:=ALLTRIM(cCodPro)
      
      IF Empty(cCodPro) 
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"-> Línea Vacia"+CRLF),NIL)
         oXls:DbSkip()   
         LOOP
      ENDIF

      IF LEN(cCodPro)>10
         IF(ValType(oSay  )="O",oSay:SetText("RIF:"+cCodPro+" tiene "+LSTR(LEN(cCodPro))+" Dígitos"),nil)
      ENDIF

      cCodPro:=LEFT(cCodPro,10)
      cFecha :=CTOO(oXls:COL_B,"C")

      IF cFecha="31/07/202"
         oXls:COL_B:=CTOD("31/07/2024")
      ENDIF

      oXls:COL_B:=CTOO(oXls:COL_B,"D")

      IF Empty(oXls:COL_B) 
         IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"-> Fecha "+cFecha+" Vacia COL_B"+CRLF),NIL)
         oXls:DbSkip()   
         LOOP
      ENDIF

      IF !ISSQLFIND("DPPROVEEDOR","PRO_RIF"+GetWhere("=",cCodPro))
         cCodPro:=EJECUTAR("DPPROVEEDORCREA",cCodPro,oXls:COL_G,cCodPro)
      ENDIF

      nContar++
      cItem     :=STRZERO(1,5)  // nContar,5) // SQLINCREMENTAL(cTable,"LBC_NUMPAR",cWhere,NIL,NIL,.T.,5)
      cNumPar   :=SQLINCREMENTAL(cTable,"LBC_NUMPAR",NIL,NIL,NIL,.T.,5)
      

      // oXls:COL_B:=CTOO(oXls:COL_B,"D")

      // #CONTROL
      oXls:COL_E:=CTOO(oXls:COL_E,"C")
      oXls:COL_D:=CTOO(oXls:COL_D,"C")

      IF Empty(oXls:COL_E)
         oXls:COL_E:=oXls:COL_D
      ENDIF

      oXls:COL_I:=CTOO(oXls:COL_I,"N") // Monto Exento
      oXls:COL_J:=CTOO(oXls:COL_J,"N") // Monto Base
      oXls:COL_K:=CTOO(oXls:COL_K,"N") // Monto IVA
      oXls:COL_L:=CTOO(oXls:COL_L,"N") // IGTG
      oXls:COL_N:=CTOO(oXls:COL_N,"C") // Cuenta Contable
      oXls:COL_O:=CTOO(oXls:COL_O,"N") // Monto pago Caja Divisa
      oXls:COL_P:=CTOO(oXls:COL_P,"N") // Monto Anticipado

      IF oDp:cTipCon="F"

        dFchDec:=EJECUTAR("GETFCHTRIMESTRE",oXls:COL_B)

/*
        IF YEAR(oXls:COL_B)<YEAR(dDesde)
          dFchDec:=CTOD("01/01/"+STRZERO(YEAR(dDesde)))
          dFchDec:=EJECUTAR("GETFCHTRIMESTRE",dFchDec)
        ELSE
          dFchDec:=EJECUTAR("GETFCHTRIMESTRE",oXls:COL_B)
        ENDIF
*/

      ENDIF

      IF !ISSQLFIND("DPCTA","CTA_CODIGO"+GetWhere("=",oXls:COL_N))
        EJECUTAR("DPCTACREA",oXls:COL_N)
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cTipDoc+"-"+ALLTRIM(oXls:COL_E)+" "+cFecha+CRLF),NIL)
      oTable:AppendBlank()
      oTable:Replace("LBC_NUMFAC",oXls:COL_D)
      oTable:Replace("LBC_NUMFIS",oXls:COL_E)
      oTable:Replace("LBC_ORIGEN","XLS")
      oTable:Replace("LBC_CODIGO",cCodPro)
      oTable:Replace("LBC_RIF"   ,cCodPro)
      oTable:Replace("LBC_FCHDEC",dFchDec   )
      oTable:Replace("LBC_FECHA" ,oXls:COL_B  )
      oTable:Replace("LBC_TIPDOC",cTipDoc     )
      oTable:Replace("LBC_CODSUC",oParXls:cCodSuc)
      oTable:Replace("LBC_USUARI",oDp:cUsuario)
      oTable:Replace("LBC_NUMPAR",cNumPar     )
      oTable:Replace("LBC_ACTIVO",.T.         )
      oTable:Replace("LBC_CXP"   ,nCxP        )
      oTable:Replace("LBC_ITEM"  ,cItem       )
      oTable:Replace("LBC_TIPIVA",cTipIva     )
      oTable:Replace("LBC_COMORG","N"         ) // Nacional
      oTable:Replace("LBC_CODMON",oDp:cMoneda ) 
      oTable:Replace("LBC_CODCTA",oXls:COL_N  ) 
      oTable:Replace("LBC_MTOEXE",oXls:COL_I  )
      oTable:Replace("LBC_BASEX" ,oXls:COL_I  )
      oTable:Replace("LBC_BASGN" ,oXls:COL_J  )
      oTable:Replace("LBC_MTOIVA",oXls:COL_K  )
      oTable:Replace("LBC_MTONET",oXls:COL_M  )
      oTable:Replace("LBC_CODMOD",oDp:cCtaMod )
      oTable:Replace("LBC_CTAEGR",cCtaEgr     )
      oTable:Replace("LBC_NUMFIL",nNumFile    )


      oTable:Replace("LBC_USOCON",oTable:LBC_CTACON)
      oTable:Replace("LBC_MTOIGT",oXls:COL_L  )
      oTable:Replace("LBC_PAGEXT",oXls:COL_O  ) // Pago en Moneda Extranjera
      oTable:Replace("LBC_ANTNAC",oXls:COL_P  ) // Anticipo en Moneda Nacional

      IF !Empty(oXls:COL_B)

         dDesde:=IF(Empty(dDesde),oXls:COL_B,dDesde)
         dDesde:=MIN(oXls:COL_B,dDesde)

         dHasta:=IF(Empty(dHasta),oXls:COL_B,dHasta)
         dHasta:=MIN(oXls:COL_B,dHasta)

      ENDIF

//      IF oXls:COL_L<>0
//        ? oXls:COL_L,"IGTF"
//      ENDIF

      oTable:SetDefault()
      oTable:Commit("")

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)

   IF(ValType(oMemo )="O",oMemo:Append("Asignando Cuentas Contables"+CRLF),NIL)
   EJECUTAR("DPLIBCOMPRASDETCTAS") // Asigna las cuentas contables desde importar datos desde excel

   IF(ValType(oMemo )="O",oMemo:Append("Importación "+LSTR(nContar)+" Registros realizada en "+LSTR(SECONDS()-nT1)+" Segundos" +CRLF),NIL)

   oTable:End(.T.)
   oXls:End()

   IF ValType(oMemo )="O"
     nT1     :=ABS(nT1-SECONDS())
     SQLUPDATE("dpfileemp",{"FIL_TEXTO","FIL_TABLE","FIL_REV"},{oMemo:GetText(),"DPLIBCOMPRASDET",LSTR(nT1)},"FIL_NUMERO"+GetWhere("=",nNumFile))
   ENDIF

   cTitle:="Libro de Compras "+DTOC(dDesde)+" "+DTOC(dHasta)

   IF !Empty(dDesde)
     EJECUTAR("BRLIBCOMEDIT",cWhere,NIL,oDp:nIndefinida,dDesde,dHasta,cTitle,CTOD(""),.F.,NIL,NIL,NIL,NIL,NIL,NIL,NIL,.F.,NIL,.F.)
   ENDIF

RETURN .T.
// EOF

