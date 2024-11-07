// Programa   : NOMINA QUINCENAL FRPREF
// Fecha/Hora : 13/04/2024
// Propósito  : Definición de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

//  AADD(aRef,{"<table>.<cField>    ,"Nombre del Campo"      ,"C"})
   AADD(aRef,{"@NOMBRE"          ,"Nombre y Apellido"        ,"C"})
   AADD(aRef,{"@RIF"             ,"Rif"                      ,"C"})
   AADD(aRef,{"@SUELDO"          ,"Quincena"                 ,"N"})
   AADD(aRef,{"@LPH"             ,"Política Habitacional"    ,"N"})
   AADD(aRef,{"@SSO"             ,"Seguro Social Obligatorio","N"})
   AADD(aRef,{"@ISLR"            ,"Impuesto Sobre la Renta"  ,"N"})
   AADD(aRef,{"@PRESTAMOS"       ,"Préstamos"                ,"N"})
   AADD(aRef,{"@BONOS"           ,"Bonos"                    ,"N"})
   AADD(aRef,{"@OTROSBONOS" 	    ,"Otros Bonos"              ,"N"})
   AADD(aRef,{"@PAGO_EFE" 	    ,"Pago en Efectivo"         ,"N"})


RETURN aRef
// EOF

