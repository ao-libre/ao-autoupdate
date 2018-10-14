# ao-autoupdate
Auto-Update to update Argentum Online client and server

![imagen](https://github.com/ao-libre/ao-autoupdate/blob/master/LEEME%20-%20Instrucciones/screenshot.jpg)

---

# '* UpdateInteligente v4.1 *
Autoupdate for any application that uses the releases system in github.

# * CONFIGURACION *

The file `ConfigAutoupdate.ini` contains the following properties

```
[ApplicationToUpdate]
application=Cliente Argentum Online Libre
repository=ao-cliente
githubAccount=ao-libre
fileToExecuteAfterUpdated=Argentum.exe
version=v13.3.13

[ConfigAutoupdate]
version=v0.1
```

ApplicationToUpdate
- application: es el nombre de la aplicacion que vamos a actualizar
- repository: es el repositorio de github de la aplicacion.
- githubAccount: ponen el nombre de usuario del repositorio que quieran utilizar
- fileToExecuteAfterUpdated: nombre del archivo que se ejecutara al finalizar la actualizacion

ConfigAutoupdate
- version: es la version del programa, para que se chequee si estamos corriendo sobre la ultima version


Errors:
En caso de tirar error en el archivo Unzip32.dll porque no se encuentra:
	Colocarla en C:/Windows/System32/ (viene adjunta al codigo)
	Ir a Inicio > Ejecutar, poner lo siguiente
	Regsvr32 Unzip32.dll
	y dar enter

-------- 

Old version / based on this version:
https://www.gs-zone.org/temas/como-configurar-el-autoupdate-completo.78653/
