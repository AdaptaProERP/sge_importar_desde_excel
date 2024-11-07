/*
// Importación de datos desde excel
// Tabla  <TABLA>
// Código condominio
// Fecha  04/03/2024
//
*/

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,oMeter,oSay,oMemo)
   LOCAL cFileDbf,cFileXls,cTable,cWhere,aTel:={},aTel2:={}
   LOCAL oTable,cNumero
   LOCAL nLinIni,nContar,oXls,cRif,cItem:="00001",cCodCli,cTipDoc,dFecha:=oDp:dFecha
   LOCAL nValCam:=oDp:nKpiValor,nCxC:=0,oDocCli,nAt
   LOCAL nT1:=SECONDS()

   DEFAULT cCodigo:="condominio"

   oTable:=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
   cFileXls:=ALLTRIM(oTable:IXL_FILE  )
   cTable  :=ALLTRIM(oTable:IXL_TABLA )
   nLinIni :=MAX(oTable:IXL_LININI,1)
   oTable:End(.T.)

   // cFileDbf:=STRTRAN( cFileXls,".xls",".dbf")
   // oXls   :=EJECUTAR("XLSTORDDFROMDEF",cCodigo)

   IF(ValType(oSay)="O",oSay:SetText("Leyendo Archivo"),NIL)

// oXls   :=EJECUTAR("XLSTORDD",cFileXls)
   oXls:=EJECUTAR("XLSTORDD",cFileXls,NIL,oMeter,oSay,NIL,nLinIni)

   IF(ValType(oMeter)="O",oMeter:SetTotal(oXls:RecCount()),NIL)

   oDocCli:=OpenTable("SELECT * FROM DPDOCCLI",.F.)

   ADEPURA(oXls:aData,{|a,n|Empty(a[1])}) // remover columnas vacias
 
   IF COUNT("DPCLIENTESREC")>0 .AND. MsgYesNo("Desea Remover todos los Registros de Propietarios")
     SQLDELETE("DPCLIENTESREC")
     SQLDELETE("DPDOCCLI","DOC_TIPDOC"+GetWhere("=","CUO"))
     SQLDELETE("DPMOVINV","MOV_TIPDOC"+GetWhere("=","CUO"))
     SQLDELETE("DPCLIENTES")
    
   ENDIF

   EJECUTAR("SETFIELDLONG","DPCLIENTES","CLI_EMAIL" ,200)

   // oXls:Browse()
   oTable:=OpenTable("SELECT * FROM DPCLIENTESREC", .F. )
   oTable:lAuditar:=.F.
   nAt:=oTable:FieldPos("CRC_USO")
   oTable:aFields[nAt,4]:=6 // 6 decimales 

   oTable:SetForeignkeyOff()

   oXls:Gotop() 

   WHILE !oXls:Eof()

      IF(ValType(oSay  )="O",oSay:SetText("Reg:"+GetNumRel(oXls:Recno(),oXls:RecCount())),NIL)
      IF(ValType(oMeter)="O",oMeter:Set(oXls:Recno()),NIL)

      cRif  :=STRTRAN(ALLTRIM(oXls:COL_C),"-","")

      IF Empty(cRif)
         cRif :=ALLTRIM(oXls:COL_A)
      ENDIF

      nAt   :=AT(";",cRif)
      IF nAt>0
         cRif:=LEFT(cRif,nAt-1)
      ENDIF

      cRif  :=STRTRAN(cRif,";","")
      cRif  :=STRTRAN(cRif,"/","")

      IF Empty(cRif)
         oXls:DbSkip()
         LOOP
      ENDIF

      IF(ValType(oMemo)="O",oMemo:Append("#"+LSTR(oXls:Recno())+"->"+cRif+CRLF),NIL)

      oXls:COL_F:=STRTRAN(oXls:COL_F,"="," ")
      oXls:COL_G:=STRTRAN(oXls:COL_G,"="," ")

      aTel  :=_VECTOR(oXls:COL_F,";")
      aTel2 :=_VECTOR(oXls:COL_G,";")

      AEVAL(aTel2,{|a,n| AADD(aTel,a)})
      ASIZE(aTel,5)

      cCodCli:=EJECUTAR("DPCLIENTECREA",cRif,oXls:COL_B,cRif) // crear Cliente

      SQLUPDATE("DPCLIENTES",{"CLI_TEL1","CLI_TEL2","CLI_TEL3","CLI_TEL4","CLI_TEL5","CLI_EMAIL"},;
                             {aTel[1]   ,aTel[2]   ,aTel[3]  ,aTel[4]    ,aTel[5]   ,oXls:COL_H },;
                             "CLI_CODIGO"+GetWhere("=",cCodCli))

      cWhere    :="CRC_CODCLI"+GetWhere("=",cRif      )+" AND "+;
                  "CRC_CODIGO"+GetWhere("=",oXls:COL_A)

      IF !ISSQLFIND("DPCLIENTESREC",cWhere)
        oTable:AppendBlank()
        oTable:Replace("CRC_ACTIVO",.T.)
        oTable:Replace("CRC_CODCLI",cCodCli)
        oTable:Replace("CRC_ID"    ,oXls:COL_A)

        IF VALTYPE(oXls:COL_D)="N"
           oTable:Replace("CRC_USO "  ,oXls:COL_D)
        ELSE
           oTable:Replace("CRC_USO "  ,CTOO(oXls:COL_D,"N"))
        ENDIF

        oTable:Replace("CRC_CENCOS",oDp:cCenCos)
        oTable:Replace("CRC_CODSUC",oDp:cSucursal)
        oTable:Replace("CRC_ITEM"  ,cItem)
       
        oTable:Commit("")

      ENDIF
      
      IF oXls:COL_J<>0
         cTipDoc:=IF(oXls:COL_J>0,"CUO","ANT")
         cNumero:=STRZERO(oXls:RecNo(),10)
         nCxC   :=IF(oXls:COL_J>0,1,-1)
         nValCam:=IF(nValCam=0,1,nValCam)
         EJECUTAR("DPDOCCLICREA",NIL,cTipDoc,cNumero,cCodCli,dFecha,oDp:cMonedaExt,"V",NIL,ABS(oXls:COL_J),0,nValCam,dFecha,NIL,oDocCli,"N",nCxC)
      ENDIF

      oXls:DbSkip()

   ENDDO

   IF(ValType(oMeter)="O",oMeter:Set(oXls:RecCount()),NIL)
   IF(ValType(oMemo )="O",oMemo:Append("Importación Concluida "+LSTR(SECONDS()-nT1)+" Segundos"+CRLF),NIL)

   oTable:End(.T.)
   oDocCli:End(.T.)
   oXls:End()
   //oXls:Browse()
RETURN .T.
