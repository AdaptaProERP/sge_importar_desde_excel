// Programa   : LIBRODECOMPRAS_YBREF
// Fecha/Hora : 09/04/2024
// Prop�sito  : Definici�n de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicaci�n :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

  AADD(aRef,{"@NOMBREPROVEEDOR"          ,"Nombre del Proveedor"          ,"C"})

RETURN aRef
// EOF

