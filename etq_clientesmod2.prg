// Programa   : CLIENTESMOD2REF
// Fecha/Hora : 21/04/2024
// Prop�sito  : Definici�n de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicaci�n :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

  AADD(aRef,{"@REPLEGAL_NOMBRE" ,"Nombre del Representante Legal" ,"C"})
  AADD(aRef,{"@REPLEGAL_RIF"    ,"Rif del Representante Legal"    ,"C"})
  AADD(aRef,{"@REPLEGAL_CORREO" ,"Correo del Representante Legal" ,"C"})
  AADD(aRef,{"@REPLEGAL_TEL"    ,"Tel�fono del Representante Legal" ,"C"})

RETURN aRef
// EOF

