
# Labs.LabEditor

 _**Hace referencia a:** aplicaciones para Office | Complementos de Office | Office Mix | PowerPoint_

El objeto **LabEditor** le permite editar un laboratorio determinado, así como obtener y establecer los datos de configuración asociados al laboratorio.

```
class LabEditor
```


## Métodos


### getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Recupera la configuración actual del laboratorio.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena una vez que se ha recuperado la configuración.|

### setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Establece una nueva configuración de laboratorio.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _configuración_|La configuración que se debe establecer.|
| _callback_|Función de devolución de llamada que se desencadena una vez que se ha establecido la configuración.|

### done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Indica que el usuario ha terminado de editar el laboratorio.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena una vez que el editor del laboratorio ha finalizado.|
