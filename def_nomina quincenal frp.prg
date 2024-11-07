/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código NOMINA QUINCENAL FRP
// Fecha  13/04/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
    LOCAL cFileDbf,cFileXls,cTable,cCodigo,cWhere
    LOCAL oTable,oXls,oTableO,oTableC
    LOCAL nLinIni,nContar,I,U
    LOCAL aRef:={},cCodPro:="",cNumero:="",cOrg:="D",nCxP:=1,aCodCta:={},cTipDoc:="NOM"

    DEFAULT cCodigo:="NOMINA QUINCENAL FRP"

    oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI,IXL_MEMO FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
    cFileXls:=ALLTRIM(oTable:IXL_FILE  )
    cTable  :=ALLTRIM(oTable:IXL_TABLA )
    nLinIni :=MAX(oTable:IXL_LININI,1)
    aRef    :=EJECUTAR("IXLLOAD",NIL,NIL,NIL,NIL,oTable:IXL_MEMO)

    AEVAL(aRef,{|a,n| aRef[n,1]:=STRTRAN(a[1],"@","")})
    oTable:End(.T.)

    // Doc CxP, Cuentas 
    SQLUPDATE("DPTIPDOCPRO",{"TDC_DOCEDI","TDC_LITEM","TDC_CONTAB"},{.T.,.T.,.T.},"TDC_TIPO"+GetWhere("=",cTipDoc))

    EJECUTAR("DPCODINTEGRA_CREA","NOMSAL","Salario")
    EJECUTAR("DPCODINTEGRA_CREA","NOMSSO","Seguro Social Obligatorio")
    EJECUTAR("DPCODINTEGRA_CREA","NOMLPH","Política Habitacional")
    EJECUTAR("DPCODINTEGRA_CREA","NOMISR","Impuesto Sobre la renta")
    EJECUTAR("DPCODINTEGRA_CREA","NOMDPR","Deducción Prestado")
    EJECUTAR("DPCODINTEGRA_CREA","NOMBPR","Bono de Producción")
    EJECUTAR("DPCODINTEGRA_CREA","NOMOBP","Otro Bono de Producción")

    aCodCta:=ASQL([SELECT CIN_CODIGO,CIC_CUENTA,CIN_DESCRI FROM dpcodintegra   LEFT JOIN DPCODINTEGRA_CTA ON CIC_CODIGO=CIN_CODIGO WHERE LEFT(CIN_CODIGO,3)="NOM"])
    AEVAL(aCodCta,{|a,n,cCodCta| cCodCta:=IF(Empty(a[2]),oDp:cCtaIndef,a[2]),;
                                 oParXls:Set("COD_"+a[1],a[1]    ),;
                                 oParXls:Set("CTA_"+a[1],cCodCta ),;
                                 oParXls:Set("DES_"+a[1],a[3]    )})

    cWhere:="DOC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
            "DOC_TIPDOC"+GetWhere("=",cTipDoc      )+" AND "+;
            GetWhereAnd("DOC_FCHDEC",oParXls:dDesde,oParXls:dHasta)

    IF COUNT(cTable,cWhere)>0 

      // .AND. MsgYesNo("Desea Remover todos los Registros de la tabla "+cTable)

      SQLDELETE("DPDOCPRO",cWhere)

      cWhere:="CCD_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
              "CCD_TIPDOC"+GetWhere("=",cTipDoc      )+" AND "+;
              GetWhereAnd("CCD_FECHA",oParXls:dDesde,oParXls:dHasta)

      SQLDELETE("DPDOCPROCTA",cWhere)

    ENDIF

    IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

    oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

    // IF(oImpXls:lBrowse .OR. .T. ,oXls:Browse(),NIL)
    // ViewArray(oXls:aData)
    // ViewArray(aRef)

    IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

    oTable:=OpenTable("SELECT * FROM "+cTable, .F. )
    oTable:lAuditar:=.F.
    oTable:SetForeignkeyOff()

    oTableC:=OpenTable("SELECT * FROM DPDOCPROCTA", .F. )
    oTableC:lAuditar:=.F.
    oTableC:SetForeignkeyOff()

    oXls:Gotop()

    WHILE !oXls:Eof()

      AEVAL(aRef,{|a,n| aRef[n,2]:=MacroEje("oXls:COL_"+a[3]),;
                        oRef:Set(aRef[n,1],aRef[n,2])})

      cCodigo:=STRTRAN(ALLTRIM(oRef:RIF),"-","")
      cCodigo:=STRTRAN(cCodigo,";","")
      cCodigo:=STRTRAN(cCodigo,"/","")

      IF !ISSQLFIND("DPPROVEEDOR","PRO_RIF"+GetWhere("=",cCodigo))
         cCodigo:=EJECUTAR("DPPROVEEDORCREA",cCodigo,oRef:NOMBRE,cCodigo)
      ELSE
         cCodigo:=SQLGET("DPPROVEEDOR","PRO_CODIGO","PRO_RIF"+GetWhere("=",cCodigo))
      ENDIF

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+cCodigo+" "+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      IF Empty(cCodigo)
         oXls:DbSkip()
         LOOP
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cCodigo+CRLF),NIL)

      cWhere:="DOC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
              "DOC_TIPDOC"+GetWhere("=",cTipDoc      )+" AND "+;
              "DOC_TIPTRA"+GetWhere("=","D"          )

      cNumero:=SQLINCREMENTAL("DPDOCPRO","DOC_NUMERO",cWhere,NIL,NIL,.T.,10)

      EJECUTAR("DPDOCPROCREA",oDp:cSucursal,cTipDoc,cNumero,"",cCodigo,oDp:dFecha,oDp:cMonedaExt,cOrg,oDp:cCenCos,oRef:SUELDO,;
                              0,oDp:nValCam,oParXls:dHasta,NIL,oTable,nCxP)

      cWhere:="DOC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
              "DOC_TIPDOC"+GetWhere("=",cTipDoc      )+" AND "+;
              "DOC_CODIGO"+GetWhere("=",cCodigo      )+" AND "+;
              "DOC_NUMERO"+GetWhere("=",cNumero      )+" AND "+;
              "DOC_TIPTRA"+GetWhere("=","D"          )

      SQLUPDATE("DPDOCPRO",{"DOC_FCHDEC"  ,"DOC_TIPORG"},;
                           {oParXls:dHasta,"XLS"       },cWhere)
	
      aCodCta:={}
      AADD(aCodCta,{oParXls:CTA_NOMBPR,oDp:cCenCos,"EX",oParXls:DES_NOMBPR,oRef:BONOS     ,0,oParXls:COD_NOMBPR,"XLS"})
      AADD(aCodCta,{oParXls:CTA_NOMDPR,oDp:cCenCos,"EX",oParXls:DES_NOMDPR,oRef:PRESTAMOS ,0,oParXls:COD_NOMDPR,"XLS"})
      AADD(aCodCta,{oParXls:CTA_NOMISR,oDp:cCenCos,"EX",oParXls:DES_NOMISR,oRef:ISLR      ,0,oParXls:COD_NOMISR,"XLS"})
      AADD(aCodCta,{oParXls:CTA_NOMLPH,oDp:cCenCos,"EX",oParXls:DES_NOMLPH,oRef:LPH       ,0,oParXls:COD_NOMLPH,"XLS"})
      AADD(aCodCta,{oParXls:CTA_NOMOBP,oDp:cCenCos,"EX",oParXls:DES_NOMOBP,oRef:OTROSBONOS,0,oParXls:COD_NOMOBP,"XLS"})
      AADD(aCodCta,{oParXls:CTA_NOMSAL,oDp:cCenCos,"EX",oParXls:DES_NOMSAL,oRef:SUELDO    ,0,oParXls:COD_NOMSAL,"XLS"})
      AADD(aCodCta,{oParXls:CTA_NOMSSO,oDp:cCenCos,"EX",oParXls:DES_NOMSSO,oRef:SSO       ,0,oParXls:COD_NOMSSO,"XLS"})

      EJECUTAR("DPDOCPROCTACREA",oDp:cSucursal,cTipDoc,cCodigo,cNumero,aCodCta,oTableC,oParXls:dHasta)

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida"+CRLF),NIL)

   oTable:End(.T.)
   oTableC:End(.T.)

   oXls:End()

RETURN .T.
