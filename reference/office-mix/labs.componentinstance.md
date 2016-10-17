
# <a name="labs.componentinstance"></a>Labs.ComponentInstance

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Representa una instancia de un componente, que es una creación de instancia de un componente dado para un usuario en tiempo de ejecución. El objeto contiene una vista traducida del componente para una ejecución específica de un laboratorio.

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## <a name="properties"></a>Propiedades

Ninguna.


## <a name="methods"></a>Métodos




### <a name="constructor"></a>Constructor

 `function constructor()`

Inicializa una nueva instancia de la clase **ComponentInstance**.


### <a name="createattempt"></a>createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

Crea un nuevo intento en el contexto de un componente.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena cuando se creó el intento.|

### <a name="getattempts"></a>getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

Recupera todos los intentos asociados al componente determinado.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena cuando se han recuperado los intentos.|

### <a name="getcreateattemptoptions"></a>getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

Recupera las opciones predeterminadas de creación de intentos. Pueden reemplazarse mediante clases derivadas.


### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

Crea un intento desde la acción determinada. Debe implementarse mediante clases derivadas.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _createAttemptResult_|La acción de creación de intentos para el intento especificado.|
