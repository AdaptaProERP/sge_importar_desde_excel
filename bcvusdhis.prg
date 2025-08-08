// Programa   : BCVUSDHIS
// Fecha/Hora : 07/08/2025 06:53:59
// Propósito  : Descargar Histórico de la Divisa desde BCV
// Creado Por : Juan Navas
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(dDesde,dHasta)
  LOCAL aTrimestre:={},cWhere,nAno1,nAno2,I,cFile,cFileXls,cFileOrg,oExcel,dFecha,cFecha
  LOCAL nLin:=15,nCol:=7,nValor:=0,cValue

  DEFAULT dDesde:=oDp:dFchInicio-500,;
          dHasta:=oDp:dFchCierre

  nAno1 :=YEAR(dDesde)
  nAno2 :=YEAR(dHasta)

  cWhere:="FCH_DESDE"+GetWhere(">=",dDesde)+" AND "+;
          "FCH_HASTA"+GetWhere("<=",dHasta)

  aTrimestre:=ASQL([ SELECT FCH_ANO,FCH_TRIMES,"" AS FILE,FCH_DESDE,FCH_HASTA ]+;
                   [ FROM VIEW_DPDIARIO_TRIMES ]+;
                   [ WHERE ]+cWhere+;
                   [ ORDER BY CONCAT(FCH_ANO,FCH_TRIMES)])

  LMKDIR("BCV")

  // https://www.bcv.org.ve/sites/default/files/EstadisticasGeneral/2_1_2b24_smc.xls

  FOR I=1 TO LEN(aTrimestre)

    cFileOrg:="2_1_2"+lower(CHR(64+aTrimestre[I,2])+RIGHT(aTrimestre[I,1],2))+"_smc.xls"
    cFileXls:=oDp:cBin+"BCV\"+cFileOrg
   
    IF !FILE(cFileXls)

       cFileOrg :="https://www.bcv.org.ve/sites/default/files/EstadisticasGeneral/"+cFileOrg

       URLDownLoad(cFileOrg,cFileXls)

    ENDIF

    IF FILE(cFileXls)
       aTrimestre[I,3]:=cFileXls
    ENDIF

  NEXT I

  SET DECI TO 4

  ADEPURA(aTrimestre,{|a,n| Empty(a[3])})

  FOR I=1 TO LEN(aTrimestre)

    cFileXls:=aTrimestre[I,3]

    oExcel:= TExcelScript():New()
    oExcel:Open( cFileXls )

    dDesde:=aTrimestre[I,4]
    dHasta:=aTrimestre[I,5]

    FOR dFecha=dHasta TO dDesde STEP -1

        cFecha:=STRTRAN(DTOC(dFecha),"/","")
        oExcel:SetSheet(cFecha) // Pestaña

        nValor:=0
        cValue:=""
        nValor:=oExcel:Get( nLin , nCol ,@cValue )

    NEXT dFecha

    oExcel:End(.F.)

  NEXT I

RETURN .t.
// EOF
