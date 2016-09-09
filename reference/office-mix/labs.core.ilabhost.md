
# Labs.Core.ILabHost

 _**Hace referencia a:** aplicaciones para Office | Complementos de Office | Office Mix | PowerPoint_

Proporciona una capa de abstracción para la conexión de Labs.js con el host.

```
interface ILabHost
```


## Métodos


### getSupportedVersions

 `getSupportedVersions(): Core.ILabHostVersionInfo[]`

Recupera las versiones admitidas por el host de laboratorio.

 **Parámetros**

Ninguno.


### connect

 `connect(versions: Core.ILabHostVersionInfo[], callback: Core.ILabCallback<Core.IConnectionResponse>)`

Inicializa una conexión con el host.

 **Parámetros**


|||
|:-----|:-----|
| _versiones_|Listado de versiones de host que el cliente puede usar.|
| _callback_|Función de devolución de llamada que se desencadena cuando la conexión se ha completado.|

### disconnect

 `disconnect(callback: Core.ILabCallback<void>)`

Finaliza la comunicación con el host.

 **Parámetros**


|||
|:-----|:-----|
| _completionStatus_|Estado del laboratorio en el momento de la desconexión.|
| _callback_|Función de devolución de llamada que se desencadena cuando se completa la desconexión.|

### on

 `on(handler: (string: any, any: any): void)`

Agrega un controlador de eventos para tratar con los mensajes que provienen del host. La promesa resuelta se devolverá de nuevo al host.

 **Parámetros**


|||
|:-----|:-----|
| _handler_|El controlador de eventos.|

### sendMessage

 `sendMessage(type: string, options: Core.IMessage, callback: Core.ILabCallback<Core.IMessageResponse>)`

Envía un mensaje al host.

 **Parámetros**


|||
|:-----|:-----|
| _tipo_|El tiempo de mensaje que se está enviando.|
| _options_|Opciones de mensaje.|
| _callback_|Función de devolución de llamada que se desencadena una vez que se recibe el mensaje.|

### create

 `create(options: Core.ILabCreationOptions, callback: Core.ILabCallback<void>)`

Crea el laboratorio. Almacena la información de host y reserva espacio para almacenar la configuración y otros elementos.

 **Parámetros**


|||
|:-----|:-----|
| _options_|Las opciones se pasan como parte de la operación de creación.|
| _callback_|Función de devolución de llamada que se desencadena una vez que se ha creado el laboratorio.|

### getConfiguration

 `getConfiguration(callback: Core.ILabCallback<Core.IConfiguration>)`

Recupera la configuración de laboratorio actual desde el host.

 **Parámetros**


|||
|:-----|:-----|
| _callback_|Función de devolución de llamada para recuperar la información de configuración.|

### setConfiguration

 `setConfiguration(configuration: Core.IConfiguration, callback: Core.ILabCallback<void>)`

Establece una nueva configuración de laboratorio en el host.

 **Parámetros**


|||
|:-----|:-----|
| _configuración_|La configuración de laboratorio que se establece.|
| _callback_|Función de devolución de llamada que se desencadena una vez que se establece la configuración.|

### getConfigurationInstance

 `getConfigurationInstance(callback: Core.ILabCallback<Core.IConfigurationInstance>)`

Recupera la configuración de instancia del laboratorio.

 **Parámetros**


|||
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena una vez que se ha recuperado la instancia de configuración.|

### getState

 `getState(callback: Core.ILabCallback<any>)`

Recupera el estado actual del laboratorio para un usuario determinado.

 **Parámetros**


|||
|:-----|:-----|
| _completionStatus_|Función de devolución de llamada que devuelve el estado de laboratorio actual.|

### setState

 `setState(state: any, callback: Core.ILabCallback<void>)`

Establece el estado del laboratorio para un usuario determinado.

 **Parámetros**


|||
|:-----|:-----|
| _state_|El estado del laboratorio.|
| _callback_|Función de devolución de llamada que se desencadena cuando se ha establecido el estado.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, callback: Core.ILabCallback<Core.IAction>)`

Realiza un intento en una acción.

 **Parámetros**


|||
|:-----|:-----|
| _tipo_|Tipo de acción.|
| _options_|Las opciones que se proporcionan con la acción.|
| _callback_|Función de devolución de llamada que devuelve la acción de ejecución final.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, result: Core.IActionResult, callback: Core.ILabCallback<Core.IAction>)`

Realiza una acción que ya se ha completado.

 **Parámetros**


|||
|:-----|:-----|
| _tipo_|Tipo de acción.|
| _options_|Las opciones que se proporcionan con la acción.|
| _result_|Resultado de la acción.|
| _callback_|Función de devolución de llamada que devuelve la acción de ejecución final.|

### getActions

 `getActions(type: string, options: Core.IGetActionOptions, callback: Core.ILabCallback<Core.IAction[]>)`

Realiza un intento en una acción.

 **Parámetros**


|||
|:-----|:-----|
| _tipo_|Tipo de obtención de acción.|
| _options_|Opciones proporcionadas con la acción de obtención.|
| _callback_|Función de devolución de llamada que devuelve el listado de las acciones completadas.|
