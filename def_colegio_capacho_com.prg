/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código COLEGIO_CAPACHO
// Fecha  25/01/2025
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cWhere,cSql,dFchAnt:=CTOD("")
    LOCAL oTable,oXls,oDb:=OpenOdbc(oDp:cDsnData),cDescri:=""
    LOCAL nLinIni,nContar,I,U,nItem:=0,nT1
    LOCAL aRef:={}
    LOCAL cItem,cNumEje,nValCam:=1,cNumero:=STRZERO(1,8),dFecha,nMonto:=0,aFechas:={},cCodCta
    LOCAL nNumFile:=NIL,dDesde:=CTOD(""),dHasta:=CTOD("")
    LOCAL cNumPar:=STRZERO(1,5),nContar:=0,cCtaCaj:="",cCtaAnt:="",cCodInt:=""

    DEFAULT cCodigo:="COLEGIO_CAPACHO_COM"


    // Cuenta contable para IGTF
    oDp:cCtaIgtfPag:=SQLGET("VIEW_DPCODINTEGRA","CIN_CODCTA,CIN_ABREVI","CIN_CODIGO='IGTFPAG'")
    oDp:cAbrIgtfPag:=ALLTRIM(DPSQLROW(2,"IGTF PAGADO"))

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_MEMO FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    aRef    :=EJECUTAR("IXLLOAD",NIL,NIL,NIL,NIL,oTable:IXL_MEMO)
    AEVAL(aRef,{|a,n| aRef[n,1]:=STRTRAN(a[1],"@","")})
    oTable:End(.T.)

    cDescri:=cFileNoPath(cFileXls)

    IF Empty(oParXls:cCodSuc)
       oParXls:cCodSuc:=oDp:cSucursal
    ENDIF

    nNumFile:=EJECUTAR("DPFILEEMPGETNUM",cFileXls,cCodigo)
    cWhere  :="MOC_NUMFIL"+GetWhere("=",nNumFile)
    nT1     :=SECONDS()

    IF nNumFile>0
      SQLDELETE(cTable,cWhere)
    ENDIF

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

   // LEER LA CUENTAS CONTABLES 
   oXls   :=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,2,4,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,.T.)

   cCtaCaj :=CTOO(oXls:K4,"C")
   cCtaAnt :=CTOO(oXls:L4,"C")

   IF !Empty(cCtaAnt)
      EJECUTAR("DPCODINTEGRA_ADD","COMANT","Anticipos Proveedores",cCtaAnt)
   ENDIF

   IF !Empty(cCtaCaj)
      EJECUTAR("DPCODINTEGRA_ADD","CAJEXT","Caja moneda Extranjera",cCtaCaj)
   ENDIF


   oXls:End()

   oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

   IF(oImpXls:lBrowse,oXls:Browse(),NIL)

   // Busca cuentas según nombre
   WHILE !oXls:Eof()

      IF Empty(oXls:COL_B) .AND. (!Empty(oXls:COL_F) .OR. !Empty(oXls:COL_G))
         cCodCta:=SQLGET("DPCTA","CTA_CODIGO","CTA_DESCRI"+GetWhere("=",oXls:COL_C))
         cCodCta:=STRTRAN(cCodCta,".","")

         IF !Empty(cCodCta)
            IF(ValType(oMemo )="O",oMemo:Append(ALLTRIM(oXls:COL_C)+" Ubicada Cta="+cCodCta+CRLF),NIL)
            oXls:Replace("COL_B",cCodCta)
         ENDIF

      ENDIF

      oXls:DbSkip()

    ENDDO

    ADEPURA(oXls:aData,{|a,n| Empty(a[2]) })

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    oTable:=OpenTable("SELECT * FROM "+cTable, .F. )

    IF Empty(oTable:aDefault)
       EJECUTAR("DPTABLESETDEF",oTable) // Facilita Asinar valores por defecto  en :AppendBlank()
    ENDIF

    cNumero:="XLS"+STRZERO(nNumFile,5) // numero del archivo de Importación XLS

    IF ValType(oMemo)="O"
       oMemo:Append("#Cbte="+cNumero+CRLF)
    ENDIF

    SET DECI TO 2

    oTable:lAuditar:=.F.
    oTable:SetForeignkeyOff()

    nContar:=0
    oXls:Gotop()

    WHILE !oXls:Eof() 
  
      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      // Asigna los valores desde oXls->oRef oReg:RIF = oXls:A
      AEVAL(aRef,{|a,n| aRef[n,2]:=MacroEje("oXls:COL_"+a[3]),;
                        oRef:Set(aRef[n,1],aRef[n,2])})


      // NO TIENE CUENTA CONTABLE COMISIONES.xlsx             
      IF Empty(oXls:COL_C) .AND. !Empty(oXls:COL_B)
         oXls:COL_C:=oXls:COL_B
         oXls:COL_B:=SQLGET("DPCTA","CTA_CODIGO","CTA_DESCRI"+GetWhere("=",oXls:COL_C))
      ENDIF

      cCodCta:=CTOO(oXls:COL_D,"C")
      cCodCta:=STRTRAN(cCodCta,"-","")
      cCodCta:=STRTRAN(cCodCta,";","")
      cCodCta:=STRTRAN(cCodCta,"/","")
      cCodCta:=STRTRAN(cCodCta,".","")

      oXls:COL_B:=CTOO(oXls:COL_B,"C")

      IF LEN(oXls:COL_B)>10
         // no es fecha
         oXls:DbSkip()
         LOOP
      ENDIF

      IF !Empty(CTOO(oXls:COL_B,"D"))
         dFecha:=CTOO(oXls:COL_B,"D")
         dDesde:=IF(Empty(dDesde),dFecha,dDesde)
         dDesde:=MIN(dDesde,dFecha)
         dHasta:=MAX(dHasta,dFecha)
         nItem :=0
      ENDIF

      IF Empty(cCodCta) .OR. Empty(dFecha)
         IF(ValType(oMemo )="O",oMemo:Append("No hay Cuenta ni Fecha "+LSTR(oXls:Recno())+CRLF),NIL)
         oXls:DbSkip()
         LOOP
      ENDIF

      // cNumero:=EJECUTAR("DPNUMCBTEXTIPDOC","DPCBTE","CBT",dFecha)
      cNumero:="XLS"+STRZERO(nNumFile,5) // numero del archivo de Importación XLS

      // dFecha :=FCHFINMES(dFecha) // solicitado por Lcdo Chona 24/01/2025

      cNumEje:=EJECUTAR("GETNUMEJE",dFecha)
      nValCam:=EJECUTAR("DPGETVALCAM",oDp:cMonedaExt,dFecha)

      cWhere:="CBT_CODSUC"+GetWhere("=",oParXls:cCodSuc)+" AND "+;
              "CBT_ACTUAL"+GetWhere("=","S"            )+" AND "+;
              "CBT_NUMERO"+GetWhere("=",cNumero        )+" AND "+;
              "CBT_FECHA" +GetWhere("=",dFecha         )

      IF dFecha<>dFchAnt .AND. !ISSQLFIND("DPCBTE",cWhere)
 
        EJECUTAR("CREATERECORD","DPCBTE",{"CBT_CODSUC"   ,"CBT_ACTUAL","CBT_NUMERO","CBT_FECHA"   ,"CBT_NUMEJE","CBT_COMEN2","CBT_ORIGEN"},;
                                         {oParXls:cCodSuc,"S"         ,cNumero     ,dFecha        ,cNumEje     ,cDescri     ,"XLS"       },;
        NIL,.T.,cWhere)

        nItem:=0

      ENDIF

      oXls:COL_C:=CTOO(oXls:COL_C,"C") // RIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodCta+CRLF),NIL)

      SET DECI TO 2

      oXls:COL_I:=IF(Empty(oXls:COL_I),0.00,oXls:COL_I)
      oXls:COL_G:=IF(Empty(oXls:COL_G),0.00,oXls:COL_G)
      oXls:COL_H:=IF(Empty(oXls:COL_H),0.00,oXls:COL_H)

      oXls:COL_F:=CTOO(oXls:COL_F,"N") // Monto Exento
      oXls:COL_G:=CTOO(oXls:COL_G,"N") // Base
      oXls:COL_H:=CTOO(oXls:COL_H,"N") // Monto IVA

      nMonto :=oXls:COL_F+oXls:COL_G+oXls:COL_H // Contribuyente formal
      // nMonto :=CTOO(oXls:COL_J,"N")

      oXls:COL_I:=CTOO(oXls:COL_I,"N") // Monto con IGTF

// IF !Empty(oXls:COL_F)
//  oXls:COL_F:=VAL(LSTR(oXls:COL_F,10,2))
//? oXls:COL_F,"oXls:COL_H",LSTR(oXls:COL_H,10,2),nMonto,"nMonto",lstr(nMonto,12,2)
// ENDIF


      // Si la cuenta no existe la busca por Nombre
      IF !ISSQLFIND("DPCTA","CTA_CODIGO"+GetWhere("=",cCodCta)) .AND. nMonto<>0
         cCodCta:=SQLGET("DPCTA","CTA_CODIGO","CTA_DESCRI"+GetWhere("=",oXls:COL_C))
         cCodCta:=STRTRAN(cCodCta,".","")
      ENDIF

      IF !ISSQLFIND("DPCTA","CTA_CODIGO"+GetWhere("=",cCodCta)) .AND. nMonto<>0

         EJECUTAR("DPCTACREA",cCodCta,oXls:COL_C)

         IF ValType(oMemo)="O"
           oMemo:Append("Cuenta "+cCodCta+" ha sido creada "+oXls:COL_C+CRLF)
         ENDIF

      ENDIF

      nMtoIva:=0 //

      IF nMonto<>0

         nContar++
         nItem++ 
         // cItem:=STRZERO(nItem,4)

         cWhere:="MOC_CODSUC"+GetWhere("=",oParXls:cCodSuc)+" AND "+;
                 "MOC_NUMCBT"+GetWhere("=",cNumero        )+" AND "+;
                 "MOC_FECHA" +GetWhere("=",dFecha         )

         cItem :=SQLINCREMENTAL("DPASIENTOS","MOC_ITEM",cWhere,NIL,NIL,.T.,4)
         oTable:AppendBlank()
         oTable:Replace("MOC_ITEM"  ,cItem)
         oTable:Replace("MOC_CUENTA",cCodCta)
         oTable:Replace("MOC_DESCRI",oXls:COL_E)
         oTable:Replace("MOC_CTAMOD",oDp:cCtaMod)
         oTable:Replace("MOC_DOCUME",""    ) // oXls:COL_B)
         oTable:Replace("MOC_FECHA" ,dFecha) // oDp:dFchInicio)
         oTable:Replace("MOC_ACTUAL","S")
         oTable:Replace("MOC_ORIGEN","XLS")
         oTable:Replace("MOC_NUMCBT",cNumero)
         oTable:Replace("MOC_NUMEJE",cNumEje)
         oTable:Replace("MOC_USUARI",oDp:cUsuario)
         oTable:Replace("MOC_MONTO" ,nMonto)
         oTable:Replace("MOC_VALCAM",nValCam)
         oTable:Replace("MOC_RIF"   ,oXls:COL_C)
         oTable:Replace("MOC_CODSUC",oParXls:cCodSuc)
         oTable:Replace("MOC_NUMFIL",nNumFile)
         oTable:Replace("MOC_NUMPAR",cNumPar )
         oTable:Replace("MOC_TIPTRA","D"     )
         oTable:Replace("MOC_TIPASI","DOC"   ) // Documento de Compras
         oTable:Replace("MOC_CODINT","COMNAC") // Compras Nacionales
         oTable:Commit("")

         IF ValType(oMemo)="O"
           DPFOCUS(oMemo)
           oMemo:Append("#"+GetNumRel(oXls:Recno(),oXls:RecCount())+" "+DTOC(dFecha)+" "+cCodCta+" "+cItem+" "+LSTR(nMonto)+CRLF)
         ENDIF

       ENDIF

         /*
         // IGTF, UTILIZO CUENTA EN MONEDA EXTRANJERA
         */
         IF oXls:COL_I>0

           cCodCta:=IF(Empty(oDp:cCtaIgtfPag),oDp:cCtaIndef,oDp:cCtaIgtfPag)
           cItem  :=SQLINCREMENTAL("DPASIENTOS","MOC_ITEM",cWhere,NIL,NIL,.T.,4)

           oTable:AppendBlank()
           oTable:Replace("MOC_ITEM"  ,cItem)
           oTable:Replace("MOC_CUENTA",cCodCta)
           oTable:Replace("MOC_DESCRI",oXls:COL_E)
           oTable:Replace("MOC_CTAMOD",oDp:cCtaMod)
           oTable:Replace("MOC_DOCUME",""    ) // oXls:COL_B)
           oTable:Replace("MOC_FECHA" ,dFecha) // oDp:dFchInicio)
           oTable:Replace("MOC_ACTUAL","S")
           oTable:Replace("MOC_ORIGEN","XLS")
           oTable:Replace("MOC_NUMCBT",cNumero)
           oTable:Replace("MOC_NUMEJE",cNumEje)
           oTable:Replace("MOC_USUARI",oDp:cUsuario)
           oTable:Replace("MOC_MONTO" ,oXls:COL_I)
           oTable:Replace("MOC_VALCAM",nValCam)
           oTable:Replace("MOC_RIF"   ,oXls:COL_C)
           oTable:Replace("MOC_CODSUC",oParXls:cCodSuc)
           oTable:Replace("MOC_NUMFIL",nNumFile)
           oTable:Replace("MOC_NUMPAR",cNumPar )
           oTable:Replace("MOC_TIPTRA","D"     )
           oTable:Replace("MOC_TIPASI","DOC"   ) // Documento de Compras
           oTable:Replace("MOC_CODINT","IGTFPAG") // Compras Nacionales
           oTable:Commit("")

         ENDIF

         // Contrapartida
         // cCtaCaj:="",cCtaAnt:=""

         nMonto :=CTOO(oXls:COL_K,"N")
        
         IF nMonto>0

           cCodCta:=cCtaCaj // Caja
           cCodInt:="CAJNAC"
           cItem  :=SQLINCREMENTAL("DPASIENTOS","MOC_ITEM",cWhere,NIL,NIL,.T.,4)

           oTable:AppendBlank()
           oTable:Replace("MOC_ITEM"  ,cItem)
           oTable:Replace("MOC_CUENTA",cCodCta)
           oTable:Replace("MOC_DESCRI",oXls:COL_E)
           oTable:Replace("MOC_CTAMOD",oDp:cCtaMod)
           oTable:Replace("MOC_DOCUME",""    ) // oXls:COL_B)
           oTable:Replace("MOC_FECHA" ,dFecha) // oDp:dFchInicio)
           oTable:Replace("MOC_ACTUAL","S")
           oTable:Replace("MOC_ORIGEN","XLS")
           oTable:Replace("MOC_NUMCBT",cNumero)
           oTable:Replace("MOC_NUMEJE",cNumEje)
           oTable:Replace("MOC_USUARI",oDp:cUsuario)
           oTable:Replace("MOC_MONTO" ,(nMonto)*-1   )
           oTable:Replace("MOC_VALCAM",nValCam)
           oTable:Replace("MOC_RIF"   ,oXls:COL_C)
           oTable:Replace("MOC_CODSUC",oParXls:cCodSuc)
           oTable:Replace("MOC_NUMFIL",nNumFile)
           oTable:Replace("MOC_NUMPAR",cNumPar )
           oTable:Replace("MOC_TIPTRA","D"     )
           oTable:Replace("MOC_TIPASI","DOC"   ) // Documento de Compras
           oTable:Replace("MOC_CODINT",cCodInt ) // Compras Nacionales
           oTable:Commit("")

         ENDIF

         nMonto :=CTOO(oXls:COL_L,"N") // Anticipo

         IF nMonto>0

           cCodCta:=cCtaAnt
           cCodInt:="COMANT"
           cItem  :=SQLINCREMENTAL("DPASIENTOS","MOC_ITEM",cWhere,NIL,NIL,.T.,4)

           oTable:AppendBlank()
           oTable:Replace("MOC_ITEM"  ,cItem)
           oTable:Replace("MOC_CUENTA",cCodCta)
           oTable:Replace("MOC_DESCRI",oXls:COL_E)
           oTable:Replace("MOC_CTAMOD",oDp:cCtaMod)
           oTable:Replace("MOC_DOCUME",""    ) // oXls:COL_B)
           oTable:Replace("MOC_FECHA" ,dFecha) // oDp:dFchInicio)
           oTable:Replace("MOC_ACTUAL","S")
           oTable:Replace("MOC_ORIGEN","XLS")
           oTable:Replace("MOC_NUMCBT",cNumero)
           oTable:Replace("MOC_NUMEJE",cNumEje)
           oTable:Replace("MOC_USUARI",oDp:cUsuario)
           oTable:Replace("MOC_MONTO" ,(nMonto)*-1   )
           oTable:Replace("MOC_VALCAM",nValCam)
           oTable:Replace("MOC_RIF"   ,oXls:COL_C)
           oTable:Replace("MOC_CODSUC",oParXls:cCodSuc)
           oTable:Replace("MOC_NUMFIL",nNumFile)
           oTable:Replace("MOC_NUMPAR",cNumPar )
           oTable:Replace("MOC_TIPTRA","D"     )
           oTable:Replace("MOC_TIPASI","DOC"   ) // Documento de Compras
           oTable:Replace("MOC_CODINT",cCodInt ) // Compras Nacionales
           oTable:Commit("")

         ENDIF

        oXls:DbSkip()

        SysRefresh(.T.)

   ENDDO

   nT1     :=ABS(nT1-SECONDS())

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida "+LSTR(nContar)+" en "+LSTR(nT1)+" segundos "+CRLF),NIL)

   IF ValType(oMemo )="O"
     SQLUPDATE("dpfileemp",{"FIL_TEXTO","FIL_TABLE","FIL_REV"},{oMemo:GetText(),"DPASIENTOS",LSTR(nT1)},"FIL_NUMERO"+GetWhere("=",nNumFile))
   ENDIF

   oTable:End(.T.)
   oXls:End()

   IF Type("oBALINIDIV")="O" .AND. oBALINIDIV:oWnd:hWnd>0
      oBALINIDIV:CLOSE()
   ENDIF

   // EJECUTAR("BRASIENTORESORG","MOC_ORIGEN"+GetWhere("=","XLS"),NIL,oDp:nIndicada,dDesde,dHasta)
   EJECUTAR("BRRECCBTCON","MOC_ORIGEN"+GetWhere("=","XLS")+" AND MOC_NUMFIL"+GetWhere("=",nNumFile),NIL,oDp:nIndicada,dDesde,dHasta)


RETURN .T.
//

