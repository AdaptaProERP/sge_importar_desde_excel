// Programa   : condominioREF
// Fecha/Hora : 26/02/2024
// Prop�sito  : Definici�n de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicaci�n :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

  AADD(aRef,{"DPCLIENTES.CLI_RIF"    ,"Rif"                 ,"C"})
  AADD(aRef,{"DPCLIENTES.CLI_NOMBRE" ,"Nombre del Cliente"  ,"C"})
  AADD(aRef,{"DPCLIENTES.CLI_TEL1"   ,"Tel�fono"            ,"C"})
  AADD(aRef,{"DPCLIENTES.CLI_TEL2"   ,"Tel�fono 2"          ,"C"})
  AADD(aRef,{"DPCLIENTES.CLI_EMAIL"  ,"Correo"              ,"C"})

RETURN aRef
// EOF

