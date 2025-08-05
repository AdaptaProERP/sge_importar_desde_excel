// Programa   : LIBRODECOMPRAS_MPREF
// Fecha/Hora : 14/04/2024
// Propósito  : Definición de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

  AADD(aRef,{"@NOMBRE","Nombre del proveedor"     ,"C"})
  AADD(aRef,{"@TIPPRO","Tipo de Proveedor"        ,"C"})
  AADD(aRef,{"@FACAFE","Factura Afectada"         ,"C"})
  AADD(aRef,{"@NUMCBT","Número de Comprobante"    ,"C"})
  AADD(aRef,{"@PLAIMP","Planilla de Importación"  ,"C"})
  AADD(aRef,{"@EXPIMP","Expediente de Importación","C"})
  AADD(aRef,{"@NUMDEB","Número Débito"            ,"C"})
  AADD(aRef,{"@NUMCRE","Número Crédito"           ,"C"})
  AADD(aRef,{"@TIPTRA","Tipo de Transacción"      ,"C"})
  AADD(aRef,{"@RIF"   ,"Rif del Proveedor"        ,"C"})


  AADD(aRef,{"@FECHA"     ,"Fecha"               ,"C"})
  AADD(aRef,{"@MTOBASIMP" ,"Monto Base Imponible","N"})
  AADD(aRef,{"@PORIVA"    ,"% IVA","N"})
  AADD(aRef,{"@MTOIVA"    ,"Monto IVA"   ,"N"})
  AADD(aRef,{"@MTOSIN"    ,"Sin derecho a Crédito Fiscal" ,"N"})
  AADD(aRef,{"@MTOEXE"    ,"Monto Exento","N"})
  AADD(aRef,{"@MTONETO"   ,"Monto Neto"  ,"N"})

  AADD(aRef,{"@MTOBAS"    ,"Monto Base Imponible","N"})
  AADD(aRef,{"@PORIVA"    ,"% IVA","N"})
  AADD(aRef,{"@MTOIVA"    ,"Monto IVA"   ,"N"})
  AADD(aRef,{"@MTOEXE"    ,"Monto Exento","N"})
  AADD(aRef,{"@MTONETO"   ,"Monto Neto"  ,"N"})

  AADD(aRef,{"@MTOBAS_IMP"    ,"Monto Base Imponible [Importación]","N"})
  AADD(aRef,{"@PORIVA_IMP"    ,"% IVA [Importación]","N"})
  AADD(aRef,{"@MTOIVA_IMP"    ,"Monto IVA [Importación]"   ,"N"})
  AADD(aRef,{"@MTOEXE_IMP"    ,"Monto Exento [Importación]","N"})
  AADD(aRef,{"@MTONETO_IMP"   ,"Monto Neto [Importación] "  ,"N"})



RETURN aRef
// EOF

