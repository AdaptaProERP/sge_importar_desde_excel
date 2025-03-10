/*
// Importaci�n de datos desde excel
// Tabla  <TABLA>
// C�digo BALACECOMPROBACION
// Fecha  30/03/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodCta,cDescri,nAt,cWhere,nMonto
    LOCAL oTable,oXls,oCta,dHasta,aFechas:={}
    LOCAL nLinIni,nContar,I,U
    LOCAL cItem,nItem:=0,cNumero:=STRZERO(1,8),cNumEje,nValCam:=1


    IF Type("oBALINIDIV")="O" .AND. oBALINIDIV:oWnd:hWnd>0
       nValCam:=oBALINIDIV:nValCam
    ENDIF
 
    DEFAULT cCodigo:="BALANCECOMPROBACION_GALAC"

    oTable:=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    oTable:End(.T.)

    SET DECI TO 2

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,1) // nLinIni)

    oXls:Goto(nLinIni-3)

    oXls:Gotop()

    WHILE !oXls:Eof() .AND. oXls:Recno()<6

      // Obtenemos la fecha final del Balance
      IF " DEL "$oXls:COL_C
         oXls:COL_C:=ALLTRIM(oXls:COL_C)
         dHasta:=CTOD(RIGHT(oXls:COL_C,10))
      ENDIF

      oXls:DbSkip()

    ENDDO

    // Fecha hasta del Balance, si esta vacio asume la fecha de inicio del ejercicio -1 , final del ejercicio pasado
    dHasta:=IF(Empty(dHasta),oDp:dFchInicio-1,dHasta)

    cNumEje:=EJECUTAR("GETNUMEJE",dHasta)
    nValCam:=EJECUTAR("DPGETVALCAM",oDp:cMonedaExt,dHasta)

    cWhere:="MOC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
            "MOC_ORIGEN"+GetWhere("=","BAL"        )+" AND "+;
            "MOC_FECHA" +GetWhere("=",dHasta       )

    IF COUNT(cTable,cWhere)>0 .AND. MsgYesNo("Desea Remover todos los Asientos del Balacen Inicial "+DTOC(dHasta))
      SQLDELETE(cTable)
    ENDIF

    cWhere:="CBT_CODSUC"+GetWhere("=",oDp:cSucursal     )+" AND "+;
            "CBT_ACTUAL"+GetWhere("=","S"               )+" AND "+;
            "CBT_NUMERO"+GetWhere("=",cNumero           )+" AND "+;
            "CBT_FECHA" +GetWhere("=",dHasta            )

    IF !ISSQLFIND("DPCBTE",cWhere)
 
       EJECUTAR("CREATERECORD","DPCBTE",{"CBT_CODSUC" ,"CBT_ACTUAL","CBT_NUMERO","CBT_FECHA"   ,"CBT_NUMEJE","CBT_COMEN1"     },;
                                        {oDp:cSucursal,"S"         ,cNumero     ,dHasta        ,cNumEje     ,"Balance Inicial"},;
       NIL,.T.,cWhere)

    ENDIF

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    oCta  :=OpenTable("SELECT * FROM DPCTA"  , .F. )
    oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
    oTable:lAuditar:=.F.
    oTable:SetForeignkeyOff() // Permite agregar registros sin integridad referencial

    oXls:Gotop()
    oXls:Goto(nLinIni) // inicia lectura

    WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cCodCta:=CTOO(oXls:COL_A,"C")

      IF "TOTAL"$UPPER(cCodCta)
         oXls:DbSkip()
         LOOP
      ENDIF

      nAt    :=AT("-",cCodCta)
      cDescri:=SUBS(cCodCta,nAt+1,LEN(cCodCta))
      cCodCta:=LEFT(cCodCta,nAt-1)
     
      // cCodCta:=STRTRAN(ALLTRIM(oXls:COL_A),"-","")
      cCodCta:=STRTRAN(cCodCta,";","")
      cCodCta:=STRTRAN(cCodCta,"/","")
      cCodCta:=STRTRAN(cCodCta,".","")
      cCodCta:=ALLTRIM(cCodCta)

      IF Empty(cCodCta) .OR. Empty(cDescri)
         oXls:DbSkip()
         LOOP
      ENDIF

      nMonto:=CTOO(oXls:COL_F,"N")

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodCta+CRLF),NIL)

      cWhere    :="CTA_CODMOD"+GetWhere("=",oDp:cCtaMod)+" AND CTA_CODIGO"+GetWhere("=",cCodCta)

      IF !ISSQLFIND("DPCTA",cWhere)
        oCta:AppendBlank()
        oCta:Replace("CTA_CODIGO",cCodCta)
        oCta:Replace("CTA_DESCRI",cDescri)
        oCta:Replace("CTA_CODMOD",oDp:cCtaMod)
        oCta:Replace("CTA_CODMOD",.T.)
        oCta:Commit("")
      ENDIF

      IF nMonto<>0

         nItem++ 
         cItem:=STRZERO(nItem,4)
         oTable:AppendBlank()
         oTable:Replace("MOC_ITEM"  ,cItem)
         oTable:Replace("MOC_CUENTA",cCodCta)
         oTable:Replace("MOC_DESCRI","Balance Inicial")
         oTable:Replace("MOC_CODMOD",oDp:cCtaMod)
         oTable:Replace("MOC_FECHA" ,dHasta) // oDp:dFchInicio)
         oTable:Replace("MOC_ACTUAL","S")
         oTable:Replace("MOC_ORIGEN","BAL")
         oTable:Replace("MOC_NUMCBT",cNumero)
         oTable:Replace("MOC_NUMEJE",cNumEje)
         oTable:Replace("MOC_USUARI",oDp:cUsuario)
         oTable:Replace("MOC_MONTO" ,nMonto)
         oTable:Replace("MOC_VALCAM",nValCam)
         oTable:Replace("MOC_CODSUC",oDp:cSucursal)
         oTable:Commit("")

      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Asignando Cuentas de Movimiento"+CRLF),NIL)
   EJECUTAR("SETDPCTADET")
   IF(ValType(oMemo )="O",oMemo:Append("Importaci�n Concluida"+CRLF),NIL)

  

   oTable:End(.T.)
   oCta:End()
   oXls:End()

   // EJECUTAR("DPCTAUSO")

   IF Type("oBALINIDIV")="O" .AND. oBALINIDIV:oWnd:hWnd>0
      oBALINIDIV:CLOSE()
   ENDIF

   IF(ValType(oMemo )="O",oMemo:Append("Creando Comprobante Inicial "+CRLF),NIL)
   EJECUTAR("DPCBTEFIX")

   aFechas:=EJECUTAR("GETFCHEJER",dHasta)
   IF !Empty(aFechas)
      EJECUTAR("BRWCOMPROBACION",NIL,aFechas[1],aFechas[2])
   ELSE
      EJECUTAR("BRWCOMPROBACION",NIL,oDp:dFchInicio,oDp:dFchCierre)
   ENDIF

RETURN .T.

