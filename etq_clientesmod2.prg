// Programa   : CLIENTESMOD2REF
// Fecha/Hora : 21/04/2024
// Propósito  : Definición de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

  AADD(aRef,{"@REPLEGAL_NOMBRE" ,"Nombre del Representante Legal" ,"C"})
  AADD(aRef,{"@REPLEGAL_RIF"    ,"Rif del Representante Legal"    ,"C"})
  AADD(aRef,{"@REPLEGAL_CORREO" ,"Correo del Representante Legal" ,"C"})
  AADD(aRef,{"@REPLEGAL_TEL"    ,"Teléfono del Representante Legal" ,"C"})

RETURN aRef
// EOF

