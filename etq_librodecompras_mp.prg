// Programa   : LIBRODECOMPRAS_MPREF
// Fecha/Hora : 14/04/2024
// Prop�sito  : Definici�n de las referencias de Campos desde Otras Tablas
// Creado Por : <AdaptaPro>
// Llamado por:
// Aplicaci�n :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL aRef:={}

  AADD(aRef,{"@NOMBRE","Nombre del proveedor"     ,"C"})
  AADD(aRef,{"@TIPPRO","Tipo de Proveedor"        ,"C"})
  AADD(aRef,{"@FACAFE","Factura Afectada"         ,"C"})
  AADD(aRef,{"@NUMCBT","N�mero de Comprobante"    ,"C"})
  AADD(aRef,{"@PLAIMP","Planilla de Importaci�n"  ,"C"})
  AADD(aRef,{"@EXPIMP","Expediente de Importaci�n","C"})
  AADD(aRef,{"@NUMDEB","N�mero D�bito"            ,"C"})
  AADD(aRef,{"@NUMCRE","N�mero Cr�dito"           ,"C"})
  AADD(aRef,{"@TIPTRA","Tipo de Transacci�n"      ,"C"})
  AADD(aRef,{"@RIF"   ,"Rif del Proveedor"        ,"C"})


  AADD(aRef,{"@FECHA"     ,"Fecha"               ,"C"})
  AADD(aRef,{"@MTOBASIMP" ,"Monto Base Imponible","N"})
  AADD(aRef,{"@PORIVA"    ,"% IVA","N"})
  AADD(aRef,{"@MTOIVA"    ,"Monto IVA"   ,"N"})
  AADD(aRef,{"@MTOSIN"    ,"Sin derecho a Cr�dito Fiscal" ,"N"})
  AADD(aRef,{"@MTOEXE"    ,"Monto Exento","N"})
  AADD(aRef,{"@MTONETO"   ,"Monto Neto"  ,"N"})

  AADD(aRef,{"@MTOBAS"    ,"Monto Base Imponible","N"})
  AADD(aRef,{"@PORIVA"    ,"% IVA","N"})
  AADD(aRef,{"@MTOIVA"    ,"Monto IVA"   ,"N"})
  AADD(aRef,{"@MTOEXE"    ,"Monto Exento","N"})
  AADD(aRef,{"@MTONETO"   ,"Monto Neto"  ,"N"})

  AADD(aRef,{"@MTOBAS_IMP"    ,"Monto Base Imponible [Importaci�n]","N"})
  AADD(aRef,{"@PORIVA_IMP"    ,"% IVA [Importaci�n]","N"})
  AADD(aRef,{"@MTOIVA_IMP"    ,"Monto IVA [Importaci�n]"   ,"N"})
  AADD(aRef,{"@MTOEXE_IMP"    ,"Monto Exento [Importaci�n]","N"})
  AADD(aRef,{"@MTONETO_IMP"   ,"Monto Neto [Importaci�n] "  ,"N"})



RETURN aRef
// EOF

