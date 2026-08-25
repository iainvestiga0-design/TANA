TANA V18 - HT DETERMINISTA, EXTRACCIÓN ESTABLE Y RESULTADOS REPRODUCIBLES

Cambios principales:
1. Gemini usa temperatura 0 en la extracción y desarrollo contable.
2. La semilla de generación se deriva del contenido exacto de la práctica. La misma práctica genera la misma semilla aunque se use otra cuenta o dispositivo.
3. Los importes de los asientos se normalizan a 2 decimales antes de construir la HT.
4. Los SALDOS AJUSTADOS de cuentas de balance se calculan como saldo neto después de aplicar los ajustes. Un ajuste parcial ya no elimina toda la cuenta.
5. ERN, ERF y ESF se alimentan de la misma HT y sus cálculos son deterministas.
6. Se mantiene la lógica contable 69<->61 y Elemento 9<->79.
7. Se mantiene el registro persistente de usuarios por correo.

Objetivo de esta versión:
La misma práctica debe producir los mismos importes y estados de forma consistente. TANA no debe compensar artificialmente un descuadre real de la práctica.


CAMBIO CLAVE V18
- La semilla de extracción y de asientos ya no depende de los metadatos binarios del archivo cuando TANA puede obtener una representación textual estable.
- DOCX y PDF con texto se normalizan localmente antes de la extracción.
- La semilla contable se calcula sobre la monografía normalizada, no sobre el archivo subido.
- La misma práctica debe producir los mismos asientos y los mismos importes aunque se suba desde otro dispositivo o correo, siempre que la extracción textual sea equivalente.
- La HT conserva los saldos parciales después de los ajustes; no elimina una cuenta completa por tener una contrapartida parcial.
- ERN, ERF y ESF siguen tomando como fuente la misma HT.
