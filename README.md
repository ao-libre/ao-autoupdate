# ao-autoupdate
Auto-Update to update Argentum Online client.

---

# '* UpdateInteligente v4.0 *

' Nuevo:

* Código reescrito y simplificado, adaptandolo a las únicas necesidades del programa
* Posibilidad de elegir que se creen los links automaticamente (EJ: http://host/Parche1.zip) o redirigir hacia un link elegido por ustedes, puede ser cualquiera (pero debe ubicarse en EJ: http://host/Link1.txt). Esto se cambia en Proyecto > Propiedades del proyecto > Generar > BuscarLinks = (0 o 1). Por defecto automático (0).
* Nueva forma de descarga de archivos más efectiva y que nos permite informar, a medida que se realiza la descarga, el tamaño del archivo descargado, su ubicacion, host y nombre.
* Nueva forma de escritura y lectura de archivos (destinado unicamente a la búsqueda del Integer del número de actualización)
* La progressbar nos indica un porcentaje preciso del tamaño del archivo
* Eliminación de elementos que quedaron en deshuso

# * CONFIGURACION *

1) Colocar el archivo Update.ini en la carpeta \INIT\ con el siguiente contenido: 0
2) Subir al host el archivo VEREXE.txt con el número de actualizaciones faltantes, 0 en caso de no haber
3) Poner los links del host en el código fuente (AutoUpdate.vbp)
4) Generar el programa
5) En caso de tirar error en el archivo Unzip32.dll porque no se encuentra:
	Colocarla en C:/Windows/System32/ (viene adjunta al codigo)
	Ir a Inicio > Ejecutar, poner lo siguiente
	Regsvr32 Unzip32.dll
	y dar enter


# * PARCHEAR *

Para colocar un parche seguir los siguientes pasos:
1) Modificar en el host el archivo http://suhost.com\VEREXE.txt incrementandole en 1 el número por cada actualización nueva
2) Subir al host los parches (Parche1.zip, Parche2.zip, etc) con los archivos a agregar al cliente (se sobreescriben y se pueden poner en carpetas). Los parches tienen que subirse con el siguiente nombre: "Parche" Numero de parche ".zip". Ejemplo "Parche2.zip"
