// Programa   : condominioREF
// Fecha/Hora : 26/02/2024
// Propósito  : Definición de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

  AADD(aRef,{"DPCLIENTES.CLI_RIF"    ,"Rif"                 ,"C"})
  AADD(aRef,{"DPCLIENTES.CLI_NOMBRE" ,"Nombre del Cliente"  ,"C"})
  AADD(aRef,{"DPCLIENTES.CLI_TEL1"   ,"Teléfono"            ,"C"})
  AADD(aRef,{"DPCLIENTES.CLI_TEL2"   ,"Teléfono 2"          ,"C"})
  AADD(aRef,{"DPCLIENTES.CLI_EMAIL"  ,"Correo"              ,"C"})

RETURN aRef
// EOF

