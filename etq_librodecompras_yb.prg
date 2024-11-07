// Programa   : LIBRODECOMPRAS_YBREF
// Fecha/Hora : 09/04/2024
// Propósito  : Definición de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

  AADD(aRef,{"@NOMBREPROVEEDOR"          ,"Nombre del Proveedor"          ,"C"})

RETURN aRef
// EOF

