# Argentum Online - Autoupdate
Auto-Update to update any application in the ecosystem of Argentum Online

![imagen](https://media.discordapp.net/attachments/496022118341935127/519694229933654017/frmLauncher.jpg)

----

# * CONFIGURATION *

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
- version: es la version del programa `ao-autoupdate`, para que se auto-chequee si estamos corriendo sobre la ultima version

---------

## Errors:
En caso de tirar error en el archivo `Unzip32.dll` porque no se encuentra:

1. Colocarla en `C:/Windows/System32/` (viene adjunta al codigo)
2. Ir a Inicio > Ejecutar, poner lo siguiente `Regsvr32 Unzip32.dll` y dar enter

-------- 

Old version / based on this version created by [@DylanUllua](https://github.com/DylanUllua) (Creator of Lhirius AO):

https://www.gs-zone.org/temas/como-configurar-el-autoupdate-completo.78653/
