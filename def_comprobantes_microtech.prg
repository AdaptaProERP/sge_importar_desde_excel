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
    LOCAL aCta:={},aCbte:={},cCbte,dFecha,nAt,nCantid,nValCam:=1,aAsientos:={},aCtas:={}
    LOCAL oDb  :=OpenOdbc(oDp:cDsnData)
    LOCAL dDesde:=CTOD(""),dHasta:=CTOD(""),nDebe,nHaber,nDec:=2,cComent:="",cTipDoc:="",cNumDoc:=""

    IF Type("oBALINIDIV")="O" .AND. oBALINIDIV:oWnd:hWnd>0
       nValCam:=oBALINIDIV:nValCam
    ENDIF
 
    DEFAULT cCodigo:="COMPROBANTES_MICROTECH"

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_LINFIN FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    nLinFin :=oTable:IXL_LINFIN

    oTable:End(.T.)

    IF nLinFin=0

       cDescri:="Introduzca el Número de la Línea Final del Archivo"+CRLF+cFileXls

       IF ValType(oMemo)="O"
         oMemo:SetText(cDescri+CRLF)
       ENDIF

       MsgMemo(cDescri)

       RETURN .F.

    ENDIF

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

    // nLinFin:=100
  
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

      IF LEFT(oXls:COL_A,12)="Comprobante:"

         oXls:COL_B:=ALLTRIM(CTOO(oXls:COL_B,"C"))
         cCbte     :=LEFT(oXls:COL_B,8)
         nAt       :=AT("Fecha:",oXls:COL_B)
         dFecha    :=CTOD(SUBS(oXls:COL_B,nAt+7,10))

         // Proxima linea es comentario
         oXls:DbSkip()
         cComent:=oXls:COL_B:=ALLTRIM(CTOO(oXls:COL_B,"C"))

         IF ValType(oMemo)="O"
            oMemo:SetText("Cbte "+cCbte+" Fecha "+DTOC(dFecha)+CRLF)
         ENDIF

         LOOP

      ENDIF

      // cCbte :=ALLTRIM(CTOO(oXls:COL_B,"C"))
     
      IF LEN(cCbte)=8 // .AND. ISALLDIGIT(cCbte)

         // dFecha :=CTOO(oXls:COL_A,"C")+"/"+SUBS(cCbte,5,2)+"/"+LEFT(cCbte,4)
         // dFecha :=CTOD(dFecha)

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
            oCbte:Replace("CBT_COMEN1",cComent)
            oCbte:Commit()

        ENDIF

      ENDIF

      oXls:COL_A:=CTOO(oXls:COL_A,"C")

      IF ISALLDIGIT(oXls:COL_A)

         cCodCta   :=CTOO(oXls:COL_B,"C")
         cCodCta   :=STRTRAN(cCodCta,";","")
         cCodCta   :=STRTRAN(cCodCta,"/","")
         cCodCta   :=STRTRAN(cCodCta,".","")
         cCodCta   :=ALLTRIM(cCodCta)

         oXls:COL_C:=CTOO(oXls:COL_C,"C")
         cDescri   :=CTOO(oXls:COL_F,"C")

         oXls:COL_G:=ALLTRIM(CTOO(oXls:COL_G,"C"))
         oXls:COL_H:=ALLTRIM(CTOO(oXls:COL_H,"C"))

         oXls:COL_G:=STRTRAN(oXls:COL_G,",",".")
         oXls:COL_H:=STRTRAN(oXls:COL_H,",",".")

         nDebe     :=VAL(oXls:COL_G)
         nHaber    :=VAL(oXls:COL_H)
         nMonto    :=IF(nDebe>0,nDebe,nHaber*-1)

         nAt       :=ASCAN(aCtas,{|a,n| a[1]==cCodCta})

         IF nAt=0
            AADD(aCtas,{cCodCta,oXls:COL_C})
         ENDIF

         oXls:COL_D:=ALLTRIM(CTOO(oXls:COL_D,"C")) // tipo
         oXls:COL_E:=ALLTRIM(CTOO(oXls:COL_E,"C")) // numero
         oXls:COL_F:=ALLTRIM(CTOO(oXls:COL_F,"C")) // Descripción del Asiento

         AADD(aAsientos,{oXls:COL_A,cCodCta,oXls:COL_F,oXls:COL_D,oXls:COL_E,nMonto,dFecha,cCbte,cNumEje})

      ENDIF
  
      oXls:DbSkip()

    ENDDO

    FOR I=1 TO LEN(aCtas)

      cCodCta:=aCtas[I,1]
      cDescri:=aCtas[I,2]

      cWhere :="CTA_CODMOD"+GetWhere("=",oDp:cCtaMod)+" AND "+;
               "CTA_CODIGO"+GetWhere("=",cCodCta    )

      IF !ISSQLFIND("DPCTA",cWhere)
         oCta:AppendBlank()
         oCta:Replace("CTA_CODIGO",cCodCta)
         oCta:Replace("CTA_DESCRI",cDescri)
         oCta:Replace("CTA_CODMOD",oDp:cCtaMod)
         oCta:Replace("CTA_ACTIVA",.T.)
         oCta:Commit("")
      ENDIF

    NEXT I

    oCta:End()
    oCbte:End()
    oXls:End()

// ViewArray(aAsientos)

    FOR I=1 TO LEN(aAsientos)

        cItem  :=aAsientos[I,1]
        cCodCta:=aAsientos[I,2]
        cDescri:=aAsientos[I,3]
        cTipDoc:=aAsientos[I,4]
        cNumDoc:=aAsientos[I,5]
        nMonto :=aAsientos[I,6]
        dFecha :=aAsientos[I,7]
        cCbte  :=aAsientos[I,8]
        cNumEje:=aAsientos[I,9]

        oAsientos:AppendBlank()
        oAsientos:Replace("MOC_ITEM"  ,cItem  )
        oAsientos:Replace("MOC_TIPO"  ,cTipDoc)
        oAsientos:Replace("MOC_DOCUME",cNumDoc)
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
    
    oAsientos:End()

    IF !Empty(dDesde)
     cWhere:="MOC_NUMFIL"+GetWhere("=",nNumFile)
     EJECUTAR("BRASIENTOSEDIT",cWhere,NIL,oDp:nIndefinida,dDesde,dHasta)
    ENDIF

RETURN .T.
// EOF
