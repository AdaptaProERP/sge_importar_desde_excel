/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código BALACECOMPROBACION
// Fecha  30/03/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodCta,cDescri,nAt,cWhere,nMonto,cFecha:=""
    LOCAL oAsientos,oXls,oCta,oCbte,oTable,dHasta,aFechas:={}
    LOCAL nLinIni,nLinFin,nContar,I,U
    LOCAL cItem,nItem:=0,cNumero:=STRZERO(1,8),cNumEje,nValCam:=1,nNumFile
    LOCAL aMes:={"ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"}
    LOCAL cMes:="",nMes:=0,cAno:=""
    LOCAL aCta:={},aCbte:={},cCbte,dFecha,nAt,nCantid,nValCam:=1,aAsientos:={}
    LOCAL oDb  :=OpenOdbc(oDp:cDsnData)
    LOCAL dDesde:=CTOD(""),dHasta:=CTOD(""),nDebe,nHaber,nDec:=2

    IF Type("oBALINIDIV")="O" .AND. oBALINIDIV:oWnd:hWnd>0
       nValCam:=oBALINIDIV:nValCam
    ENDIF
 
    DEFAULT cCodigo:="MAYOR_ANALITICO_MICROTECH"

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_LINFIN FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    nLinFin :=MAX(oTable:IXL_LINFIN,0)


    oTable:End(.T.)

//    SET DECI TO 2

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    nNumFile:=EJECUTAR("DPFILEEMPGETNUM",cFileXls,cCodigo)

    cWhere  :="CBT_NUMFIL"+GetWhere("=",nNumFile)
    nCantid :=COUNT("DPCBTE",cWhere)

    IF nCantid>0 

      SQLDELETE("DPCBTE",cWhere)

      cWhere  :="MOC_NUMFIL"+GetWhere("=",nNumFile)

      SQLDELETE("DPASIENTOS",cWhere)

    ENDIF

    nLinFin:=100
  
    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni,nLinFin,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,nDec) // nLinIni)

    SET DECI TO (nDec)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    // oCta  :=OpenTable("SELECT * FROM DPCTA"  , .F. )
    // oCta  :lAuditar:=.F.
    // oCta  :SetForeignkeyOff()

    oCta     :=TINSERTINTO("DPCTA"     ,oDb,15)
    oAsientos:=TINSERTINTO("DPASIENTOS",oDb,100)
    oCbte    :=TINSERTINTO("DPCBTE"    ,oDb,15)

    oXls:Gotop()

    WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      IF LEFT(oXls:COL_A,6)="Cuenta"

        cCodCta:=CTOO(SUBS(oXls:COL_A,8,20),"C")
        cCodCta:=STRTRAN(cCodCta,";","")
        cCodCta:=STRTRAN(cCodCta,"/","")
        cCodCta:=STRTRAN(cCodCta,".","")
        cCodCta:=ALLTRIM(cCodCta)
        cDescri:=CTOO(oXls:COL_F,"C")

        cWhere :="CTA_CODMOD"+GetWhere("=",oDp:cCtaMod)+" AND "+;
                 "CTA_CODIGO"+GetWhere("=",cCodCta    )

        IF !ISSQLFIND("DPCTA",cWhere)
          oCta:AppendBlank()
          // oCta:lAuditar:=.F.
          oCta:Replace("CTA_CODIGO",cCodCta)
          oCta:Replace("CTA_DESCRI",cDescri)
          oCta:Replace("CTA_CODMOD",oDp:cCtaMod)
          oCta:Replace("CTA_ACTIVA",.T.)
          oCta:Commit("")
        ENDIF

      ENDDI

      cCbte :=ALLTRIM(CTOO(oXls:COL_B,"C"))
     
      IF LEN(cCbte)=8 .AND. ISALLDIGIT(cCbte)

         dFecha :=CTOO(oXls:COL_A,"C")+"/"+SUBS(cCbte,5,2)+"/"+LEFT(cCbte,4)
         dFecha :=CTOD(dFecha)

         dDesde :=IF(Empty(dDesde),dFecha,dDesde)
         dHasta :=IF(Empty(dHasta),dFecha,dHasta)

         dDesde :=MIN(dDesde,dFecha)
         dHasta :=MAX(dHasta,dFecha)

         nAt    :=ASCAN(aCbte,{|a,n| a[1]==cCbte .AND. a[2]=dFecha})
         cNumEje:=EJECUTAR("GETNUMEJE",dFecha)

         IF nAt=0 
         
            AADD(aCbte,{cCbte,dFecha})

            oCbte:AppendBlank()
            oCbte:SetDefault()
            // oCbte:lAuditar:=.F.
            oCbte:Replace("CBT_NUMERO",cCbte)
            oCbte:Replace("CBT_FECHA" ,dFecha)
            oCbte:Replace("CBT_CODSUC",oDp:cSucursal)
            oCbte:Replace("CBT_NUMEJE",cNumEje)
            oCbte:Replace("CBT_ACTUAL","S")
            oCbte:Replace("CBT_NUMFIL",nNumFile)
            oCbte:Commit()

        ENDIF

        oXls:COL_F:=CTOO(oXls:COL_F,"C")
        oXls:COL_G:=ALLTRIM(CTOO(oXls:COL_G,"C"))
        oXls:COL_H:=ALLTRIM(CTOO(oXls:COL_H,"C"))

        oXls:COL_G:=STRTRAN(oXls:COL_G,",",".")
        oXls:COL_H:=STRTRAN(oXls:COL_H,",",".")

        nDebe  :=VAL(oXls:COL_G)
        nHaber :=VAL(oXls:COL_H)

        nMonto :=IF(nDebe>0,nDebe,nHaber*-1)
        cDescri:=CTOO(oXls:COL_F,"C")

        AADD(aAsientos,{cCodCta,nMonto,dFecha,cDescri,cCbte,cNumEje})

      ENDIF
   
      oXls:DbSkip()

    ENDDO

    oCta:End()
    oCbte:End()

    FOR I=1 TO LEN(aAsientos)

        cCodCta:=aAsientos[I,1]
        nMonto :=aAsientos[I,2]
        dFecha :=aAsientos[I,3]
        cDescri:=aAsientos[I,4]
        cCbte  :=aAsientos[I,5]
        cNumEje:=aAsientos[I,6]

        cItem:="0000" // STRZERO(nItem,4)
        oAsientos:AppendBlank()
        oAsientos:Replace("MOC_ITEM"  ,cItem)
        oAsientos:Replace("MOC_CUENTA",cCodCta)
        oAsientos:Replace("MOC_DESCRI",cDescri)
        oAsientos:Replace("MOC_CTAMOD",oDp:cCtaMod)
        oAsientos:Replace("MOC_FECHA" ,dFecha ) // oDp:dFchInicio)
        oAsientos:Replace("MOC_ACTUAL","S"    )
        oAsientos:Replace("MOC_ORIGEN","XLS"  )
        oAsientos:Replace("MOC_NUMCBT",cCbte  )
        oAsientos:Replace("MOC_NUMEJE",cNumEje)
        oAsientos:Replace("MOC_USUARI",oDp:cUsuario)
        oAsientos:Replace("MOC_MONTO" ,nMonto)
        oAsientos:Replace("MOC_VALCAM",nValCam)
        oAsientos:Replace("MOC_CODSUC",oDp:cSucursal)
        oAsientos:Replace("MOC_NUMFIL",nNumFile)
        oAsientos:Commit("")

    NEXT I

    oXls:End()
    
    oAsientos:End()
 
//    EJECUTAR("BRRECCBTCON","MOC_ORIGEN"+GetWhere("=","XLS")+" AND MOC_NUMFIL"+GetWhere("=",nNumFile),NIL,oDp:nIndicada,dDesde,dHasta)

    EJECUTAR("BRWMAYORANALITICO",NIL,dDesde,dHasta)
//
// PROCE MAIN(oGenRep,dDesde,dHasta,RGO_C1,RGO_C2,RGO_C3,RGO_C4,RGO_I1,RGO_F1,RGO_I2,RGO_F2,cCodMon,nPeriodo,cWhereAdd)


RETURN NIL
// EOF
