// Programa   : DPIMPRXLSINV
// Fecha/Hora : 11/03/2014 02:34:15
// Propósito  : Importar Productos desde EXCEL
// Creado Por : Juan Navas
// Llamado por: DPIMPRXLSRUN   
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodDef,lChk,lTodos,nCantid,oMemo,oMeterR,oSay,lBrowse)
   LOCAL cFileDbf,cFileXls,nLinIni
   LOCAL oInv,cGrupo,cMarca,cWhere,cLine,lOk,oEqui,cBarra,cMemoX:="",oSuc,oTable
   LOCAL aSelect:={},aCampos:={}
   LOCAL cField,nAt,n,aVar:={},cVar,cGruNombre:="",cMarNombre:="",oTable
   LOCAL cPrecio_A:="",cPrecio_B:="",cPrecio_C:="",cPrecio_D:="",cPrecio_E:="",cPrecio_L:=""
   LOCAL cUndMed:="",cPeso:="",cCXUNDMED:="",cPRESENTA:=""
   LOCAL cCANT  :="",cCOSTO:="",cLOTE:="",cFCHVENC:="",cCodSuc:="",cCodAlm:=""
   LOCAL cLIMSUC:="SI",cWhereS:="",oTabSuc
   LOCAL n_:=0,cMemo:="",cFileIxl:="",cTable,cCodigo

   DEFAULT lChk   :=.T.,;
           lTodos :=.F.,;
           nCantid:=1  ,;
           lBrowse:=.T.

   DEFAULT cCodDef:=SQLGET("DPIMPRXLS","IXL_CODIGO")

   oTable  :=OpenTable("SELECT * FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodDef),.T.)
   cFileXls:=ALLTRIM(oTable:IXL_FILE  )
   cTable  :=ALLTRIM(oTable:IXL_TABLA )
   cMemo   :=ALLTRIM(oTable:IXL_MEMO  )
   nLinIni :=MAX(oTable:IXL_LININI,1)

? cFileXls,"file"

   cfileDbf:=STRTRAN( cFileXls,".xls",".dbf")

   IF !FILE(ALLTRIM(cFileXls))
      MensajeErr("Archivo "+cFileXls+" no Existe")
      RETURN NIL
   ENDIF

   cFileIxl:=cTable

   aSelect:=ACLONE(EJECUTAR("IXLLOAD",cTable,NIL,cCodDef))

   IF Empty(aSelect)
      MensajeErr("No hay definición de Campos para importar")
      EVAL(oBtnCnf:bAction)
      RETURN NIL
   ENDIF

   IF Empty(EJECUTAR("XLSTODBF",cFileXls,cFileDbf,oMeterR,oSay,.T.,nLinIni,IF(lTodos,0,nCantid)))
      RETURN .T.
   ENDIF

   IIF(oSay=NIL,NIL,oSay:SetText("Leyendo Datos desde "+cFileDbf))

   CLOSE ALL

   SELE A
   USE (cFileDbf) SHARED

   IF RECCOUNT()=0
      MensajeErr("No se realizó la Lectura de Registros, Revise Número de Línea de Lectura")
      CLOSE ALL
      nLinIni++
//    oLinIni:VarPut(nLinIni,.T.) 
      RETURN .F.
   ENDIF


   IF lBrowse
     BROWSE()
     CLOSE ALL
     RETURN NIL
   ENDIF

   cMemoX:="Iniciando Importación"+CRLF+"Leyenda "+CRLF+"I:Incluir M:Modificar B:Barra"

   IIF(oMemo=NIL,NIL,oMemo:VarPut(cMemoX,.T.))

   IF Empty(ALIAS()) 
      oMemo:Append("Archivo "+cFileDbf)
      RETURN .F.
   ENDIF

   IF !oMeterR=NIL
     oMeterR:SetTotal(RECCO())
   ENDUF

   aCampos:=ACLONE(DBSTRUCT())

   AEVAL(aCampos,{|a,n|aCampos[n]  :=ALLTRIM(a[1])})
   AEVAL(aSelect,{|a,n|aSelect[n,3]:=ALLTRIM(a[3])})

   FOR n_=1 TO LEN(aSelect)

      nAt:=0

      IF !Empty(aSelect[n_,3])
         nAt:=ASCAN(aCampos,aSelect[n_,3])
      ENDIF

      IF nAt>0
         nAt:=FIELDPOS(aCampos[nAt])
      ENDIF

      aSelect[n_,2]:=nAt

      IF "@"$aSelect[n_,1]
        cVar:="c"+STRTRAN(aSelect[n_,1],"@","")
        PUBLICO(cVar)
        AADD(aVar,cVar)
      ENDIF

   NEXT n_

   SETEXCLUYE("DPGRU"        , "")
   SETEXCLUYE("DPMARCAS"     , "")
   SETEXCLUYE("DPINV"        , "")
   SETEXCLUYE("DPPRECIOTIP"  , "")
   SETEXCLUYE("DPPRECIOS"    , "")

// oDp:lTracer:=.T.

   WHILE !A->(EOF()) 

  
      cLine  :=""

      IF !oMererR=NIL
        oMeterR:Set(RECNO())  
      ENDIF

      IIF(oSay=NIL,NIL,oSay:SetText(LSTR(RECCO())+"/"+LSTR(RECNO())))

      // Asigna los Valores en el Objeto
      oInv :=OpenTable("SELECT * FROM DPINV",.F.)

      FOR n_=1 TO LEN(aSelect)

         IF "@"$aSelect[n_,1]

            cVar:="c"+STRTRAN(aSelect[n_,1],"@","")

            IF aSelect[n_,2]>0

               PUBLICO(cVar,FIELDGET(aSelect[n_,2]))

            ELSE

               PUBLICO(cVar,aSelect[n_,4])

            ENDIF
           
         ELSE

           IF aSelect[n_,2]>0
              oInv:Replace(aSelect[n_,1],FIELDGET(aSelect[n_,2]))
           ENDIF

           // Valor por defecto
           IF !Empty(aSelect[n_,4]) .AND. Empty(oInv:Get(aSelect[n_,1]))
             oInv:Replace(aSelect[n_,1],aSelect[n_,4])
           ENDIF

         ENDIF

      NEXT 

      cCodigo  :=oInv:INV_CODIGO
      cCodigo  :=CTOO(cCodigo,"C")
      cBarra   :=IIF(cBarra=NIL,"",cBarra)
      cBarra   :=CTOO(cBarra   ,"C")
      cUndMed  :=CTOO(cUndMed  ,"C")
      cPRESENTA:=CTOO(cPRESENTA,"C")

      IF ValType(cPeso)="C"
         cPeso    :=VAL(cPeso)
      ENDIF

      cCXUNDMED:=CTOO(cCXUNDMED,"N")

      IF Empty(cUndMed)
         cUndMed:="UND"
      ENDIF

      IF Empty(cCXUNDMED)
         cCXUNDMED:=1
      ENDIF

    
      // Busca si Existe la Unidad de Medida

      IF !lChk
        CreaUndMed(cUndMed,cPeso)
      ENDIF

      cWhere:="INV_CODIGO"+GetWhere("=",cCodigo)

      cGrupo:=GetGrupo(oInv:INV_GRUPO ,cGruNombre)
      cMarca:=GetMarca(oInv:INV_CODMAR,cMarNombre)

      IF Empty(cGrupo)
         cGrupo:=GetGrupo("INDEF","Indefinido")
      ENDIF

      IF Empty(cMarca)
        cMarca:=GetMarca("INDEF","Indefinido")
      ENDIF

      oInv:Replace("INV_GRUPO" ,cGrupo)
      oInv:Replace("INV_CODMAR",cMarca)

      IF Empty(oInv:INV_CODMAR)
         cLine:="Codigo "+cCodigo+" Marca Vacia"
      ENDIF

      IF Empty(oInv:INV_GRUPO)
         cLine:="Codigo "+cCodigo+" Grupo Vacio"
      ENDIF

      IF Empty(oInv:INV_IVA)
        oInv:Replace("INV_IVA","GN")
      ENDIF

      IF Empty(oInv:INV_UTILIZ)
        oInv:Replace("INV_UTILIZ","V")
      ENDIF

      IF Empty(oInv:INV_ESTADO)
        oInv:Replace("INV_ESTADO","A")
      ENDIF

      IF Empty(oInv:INV_APLICA)
        oInv:Replace("INV_APLICA","T")
      ENDIF

      IF Empty(oInv:INV_PROCED)
        oInv:Replace("INV_PROCED","N")
      ENDIF

      IF Empty(oInv:INV_METCOS)
        oInv:Replace("INV_METCOS","P")
      ENDIF

      IF !Empty(cLote)
        oInv:Replace("INV_METCOS","C")
      ENDIF

      IF !Empty(cCodSuc)
         cCodSuc:=CTOO(cCodSuc,"N")
         cCodSuc:=STRZERO(cCodSuc,6)
      ELSE
         cCodSuc:=oDp:cSucursal
      ENDIF

      IF oDp:lInvXSuc 

         EJECUTAR("DPINVSETSUC",cCodigo,cCodSuc)
/*
         oTabSuc:=OpenTable(" SELECT * FROM DPTABXSUC WHERE "           +;
                            " RXS_CLAVE "+GetWhere("=",cCodigo )+" AND "+;
                            " RXS_TABLA "+GetWhere("=","DPINV" )+" AND "+;
                            " RXS_CODSUC"+GetWhere("=",cCodSuc ),.T.)
                   
         IF oTabSuc:RecCount()=0
            oTabSuc:Append()
            cWhereS:=""
         ELSE
            cWhereS:=oTabSuc:cWhere
         ENDIF

         oTabSuc:Replace("RXS_CODSUC",cCodSuc        )
         oTabSuc:Replace("RXS_CLAVE" ,cCodigo        )
         oTabSuc:Replace("RXS_TABLA" ,"DPINV"        )
         oTabSuc:Replace("RXS_SELECT",.T.)
         oTabSuc:Commit(cWhereS)
         oTabSuc:End()
*/

     ENDIF


      IF Empty(SQLGET("DPSUCURSAL","SUC_CODIGO","SUC_CODIGO"+GetWhere("=",cCodSuc)))
        oSuc:=OpenTable("DPSUCURSAL",.F.)
        oSuc:AppendBlank()
        oSuc:Replace("SUC_CODIGO",cCodSuc)
        oSuc:Replace("SUC_DESCRI","Creada desde Importar desde EXCEL")
        oSuc:Commit()
      ENDIF

      // Si desea Intercambiar Codigo x Barra, quitar comentarios
      // cCodigo:=CTOO(A->E,"C")
      // cBarra :=CTOO(A->C,"C")
      // es necesario Borrar los productos y Barras
      // DELETE FROM DPEQUIV;
      // DELETE FROM DPINV

      IF Empty(cCodigo)
         cline:="Código Vacio"
      ENDIF

      IF !Empty(cLine)

         lOk:=.F.

         IIF(oMemo=NIL,NIL, oMemo:Append(cLine+CRLF))
         cMemoX  :=cMemoX+IF(Empty(cLine),CRLF,"")+cLine

         A->(DbSkip())
         LOOP

      ENDIF

      cWhere:="INV_CODIGO"+GetWhere("=",cCodigo)


      oTable:=OpenTable("SELECT * FROM DPINV WHERE "+cWhere,.T.)

      cLine:=CRLF+IIF(oTable:RecCount()=0,"I:","M:")+ALLTRIM(cCodigo)+;
             IF(Empty(cBarra)," "," B:" +ALLTRIM(cBarra))

      IF oTable:RecCount()=0
   
         cWhere:=""
         oTable:AppendBlank()

      ENDIF

      IF(oMemo=NIL,NIL,oMemo:Append(cLine))
      cMemoX  :=cMemoX+cLine

      AEVAL(oTable:aFields,{|a,n| oTable:FieldPut(n,oInv:FieldGet(n)) })

      oTable:Replace("INV_CODIGO",cCodigo)
      oTable:Replace("INV_APLORG","XLS"  )  
      
      // Si esta en Revision, No agrega producto

      IF !lChk
         oTable:Commit(cWhere)
         cPeso:=IF(Empty(cPeso),NIL,cPeso) // Tomara el peso de la Unidad de Medida
         EJECUTAR("DPINVCREAUND",cCodigo,cUndMed,cCXUNDMED,cPeso,cPRESENTA)
      ENDIF

      oTable:End()

      oInv:End()

      // Equivalentes
      IF !Empty(cBarra) .AND. !lChk 

        cWhere:="EQUI_BARRA"+GetWhere("=",cBarra )+" AND "+;
                "EQUI_CODIG"+GetWhere("=",cCodigo)

        oEqui:=OpenTable("SELECT * FROM DPEQUIV WHERE "+cWhere,.T.)

        IF oEqui:RecCount()=0
           oEqui:AppendBlank()
        ENDIF
       
        oEqui:Replace("EQUI_BARRA",cBarra )
        oEqui:Replace("EQUI_CODIG",cCodigo)
        oEqui:Replace("EQUI_MED"  ,cUndMed)

        oEqui:Commit(cWhere)
        oEqui:End()

      ENDIF

      IF !Empty(cPrecio_A) .AND. !lChk 
         CREAR_PRECIO(cCodigo,"A",cPrecio_A,cUndMed)
      ENDIF

      IF !Empty(cPrecio_B) .AND. !lChk 
         CREAR_PRECIO(cCodigo,"B",cPrecio_B,cUndMed)
      ENDIF

      IF !Empty(cPrecio_C) .AND. !lChk 
         CREAR_PRECIO(cCodigo,"C",cPrecio_C,cUndMed)
      ENDIF

      IF !Empty(cPrecio_D) .AND. !lChk 
         CREAR_PRECIO(cCodigo,"D",cPrecio_D,cUndMed)
      ENDIF

      IF !Empty(cPrecio_E) .AND. !lChk 
         CREAR_PRECIO(cCodigo,"E",cPrecio_E,cUndMed)
      ENDIF

      IF !Empty(cCosto) .AND. !Empty(cCant)

          cCosto :=CTOO(cCosto,"N")
          cCant  :=CTOO(cCant ,"N")
          cCodAlm:=IF(Empty(cCodAlm),oDp:cAlmacen,cCodAlm)
          cCodAlm:=CREAALMACEN(cCodAlm)

          IF ValType(cFchVenc)="C" .AND. LEN(ALLTRIM(cFCHVENC))=7
            cFCHVENC:="01/"+cFCHVENC
          ENDIF

          cFCHVENC:=CTOD(cFCHVENC)

          EJECUTAR("DPINVEXIINI",cCodigo,cCant,cCosto,cUndMed,NIL,cLote,cFCHVENC,cPrecio_L,cCodSuc,cCodAlm)

      ENDIF

      A->(DbSkip())

// IF RECNO()>5
//   EXIT
// ENDIF

   ENDDO

   IF(oMemo=NIL,NIL,oMemo:Append(CRLF+"Proceso Concluido"))
 
   EJECUTAR("DPBUILDWHERE")

   CLOSE ALL

   EJECUTAR("AUDITORIA","PROC",NIL,"DPIMPRXLS",cCodigo)
    
RETURN lOk

/*
// Obtiene Grupo
*/
FUNCTION GETGRUPO(cGrupo,cDescri)
  LOCAL oTable

  IF Empty(cGrupo) .AND. !Empty(cDescri)
     cGrupo:=SQLGET("DPGRU","GRU_CODIGO","GRU_DESCRI"+GetWhere("=",cDescri))
  ENDIF

  IF Empty(cGrupo) .AND. Empty(cDescri)
     cGrupo:=SQLGET("DPGRU","GRU_CODIGO")
  ENDIF

  IF Empty(cGrupo)
     cGrupo:=SQLINCREMENTAL("DPGRU","GRU_CODIGO")
  ENDIF

  cGrupo:=ALLTRIM(cGrupo)

  IF ALLTRIM(SQLGET("DPGRU","GRU_CODIGO","GRU_CODIGO"+GetWhere("=",cGrupo)))=cGrupo
     RETURN cGrupo
  ENDIF

  oTable:=OpenTable("SELECT * FROM DPGRU",.F.)
  oTable:Append()
  oTable:Replace("GRU_CODIGO",cGrupo )
  oTable:Replace("GRU_DESCRI",cDescri)
  oTable:Replace("GRU_ACTIVO",.T.    )
  oTable:Commit()
  oTable:End()

RETURN cGrupo

/*
// Obtiene MARCA
*/
FUNCTION GETMARCA(cMarca,cDescri)
  LOCAL oTable

  IF Empty(cMarca) .AND. !Empty(cDescri)
     cMarca:=SQLGET("DPMARCAS","MAR_CODIGO","MAR_DESCRI"+GetWhere("=",cDescri))
  ENDIF

  IF Empty(cMarca) .AND. Empty(cDescri)
     cMarca:=SQLGET("DPMARCAS","MAR_CODIGO")
  ENDIF

  IF Empty(cMarca)
    cMarca:=SQLINCREMENTAL("DPMARCAS","MAR_CODIGO")
  ENDIF

  cMarca:=ALLTRIM(cMarca)

  IF ALLTRIM(SQLGET("DPMARCAS","MAR_CODIGO","MAR_CODIGO"+GetWhere("=",cMarca)))=cMarca
     RETURN cMarca
  ENDIF

  oTable:=OpenTable("SELECT * FROM DPMARCAS",.F.)
  oTable:Append()
  oTable:Replace("MAR_CODIGO",cMarca )
  oTable:Replace("MAR_DESCRI",cDescri)
  oTable:Commit()
  oTable:End()

RETURN cMarca

FUNCTION CONFIGCOL()
   LOCAL aCol:={}

   AADD(aCol,{"INV_CODIGO","","A"})
   AADD(aCol,{"INV_DESCRI","","B"})

RETURN NIL

FUNCTION CREAR_PRECIO(cCodigo,cLista,nPrecio,cUndMed)
   LOCAL cWhere,oPrecio
  
   DEFAULT cUndMed:=oDp:cUndMed

   nPrecio:=CTOO(nPrecio,"N")

   cWhere:="PRE_CODIGO"+GetWhere("=",cCodigo     )+ " AND "+;
           "PRE_UNDMED"+GetWhere("=",cUndMed     )+ " AND "+;
           "PRE_LISTA" +GetWhere("=",cLista      )+ " AND "+;
           "PRE_CODMON"+GetWhere("=",oDp:cMoneda )


   oPrecio:=OpenTable("SELECT * FROM DPPRECIOS WHERE  "+cWhere,.T.)

   IF oPrecio:RecCount()=0
      oPrecio:AppendBlank()
      cWhere :=""
   ENDIF

   oPrecio:Replace("PRE_CODIGO",cCodigo        )
   oPrecio:Replace("PRE_UNDMED",cUndMed        )
   oPrecio:Replace("PRE_LISTA" ,cLista         )
   oPrecio:Replace("PRE_PRECIO",nPrecio        )
   oPrecio:Replace("PRE_CODMON",oDp:cMoneda    )
   oPrecio:Replace("PRE_FECHA" ,oDp:dFecha     )
   oPrecio:Replace("PRE_HORA"  ,TIME()         )
   oPrecio:Replace("PRE_USUARI",oDp:cUsuario   )
   oPrecio:Replace("PRE_ORIGEN","F"            )
   oPrecio:Replace("PRE_IP"    ,GETHOSTBYNAME())

   oPrecio:Commit(cWhere)
   oPrecio:End()

RETURN NIL


FUNCTION CREAUNDMED(cUndMed,cPeso)
  LOCAL oTable

  IF !Empty(SQLGET("DPUNDMED","UND_CODIGO","UND_CODIGO"+GetWhere("=",cUndMed)))
     RETURN cUndMed
  ENDIF

  oTable:=OpenTable("SELECT * FROM DPUNDMED",.F.)
  oTable:Append()
  oTable:Replace("UND_CODIGO",cUndMed )
  oTable:Replace("UND_DESCRI",cUndMed )
  oTable:Commit()
  oTable:End()

RETURN cUndMed

FUNCTION CREAALMACEN(cCodAlm)
  LOCAL oTable

  IF !Empty(SQLGET("DPALMACEN","ALM_CODIGO","ALM_CODIGO"+GetWhere("=",cCodAlm)))
     RETURN cCodAlm
  ENDIF

  oTable:=OpenTable("SELECT * FROM DPALMACEN",.F.)
  oTable:Append()
  oTable:Replace("ALM_CODIGO",cCodAlm )
  oTable:Replace("ALM_DESCRI","Creado desde "+cFileNoPath(cFileXls))
  oTable:Commit()
  oTable:End()

RETURN cCodAlm

// EOF
