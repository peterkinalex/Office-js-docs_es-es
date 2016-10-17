
# <a name="labs.labinstance"></a>Labs.LabInstance

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Una instancia de un laboratorio que está configurado para el usuario actual. Use este objeto para grabar y recuperar datos de laboratorio para el usuario.

```
class LabInstance
```


## <a name="variables"></a>Variables


|||
|:-----|:-----|
| `public var data: any`|Variable de contenedor para almacenar los datos de usuario.|
| `public var components: Labs.ComponentInstanceBase[]`|Los componentes que crean la instancia de laboratorio.|

## <a name="methods"></a>Métodos




### <a name="getstate"></a>getState

 `public function getState(callback: Labs.Core.ILabCallback<any>): void`

Recupera el estado actual del laboratorio para un usuario determinado.

 **Parámetros**


|||
|:-----|:-----|
| _callback_|La función de devolución de llamada que se desencadena cuando se recupera el estado del laboratorio.|

### <a name="setstate"></a>setState

 `public function setState(state: any, callback: Labs.Core.ILabCallback<void>): void`

Establece el estado del laboratorio para un usuario determinado.

 **Parámetros**


|||
|:-----|:-----|
| _state_|Estado que se debe establecer.|
| _callback_|Función de devolución de llamada que se desencadena una vez que se establece el estado.|

### <a name="done"></a>Done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Función de indicador que indica que el usuario ha finalizado la realización del laboratorio.

 **Parámetros**


|||
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena una vez que el laboratorio ha finalizado.|
