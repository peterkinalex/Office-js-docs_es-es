
# Labs.ComponentInstance

 _**Hace referencia a:** aplicaciones para Office | Complementos de Office | Office Mix | PowerPoint_

Representa una instancia de un componente, que es una creación de instancia de un componente dado para un usuario en tiempo de ejecución. El objeto contiene una vista traducida del componente para una ejecución específica de un laboratorio.

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## Propiedades

Ninguna.


## Métodos




### Constructor

 `function constructor()`

Inicializa una nueva instancia de la clase **ComponentInstance**.


### createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

Crea un nuevo intento en el contexto de un componente.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena cuando se ha creado el intento.|

### getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

Recupera todos los intentos asociados al componente determinado.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena cuando se han recuperado los intentos.|

### getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

Recupera las opciones predeterminadas de creación de intentos. Pueden reemplazarse mediante clases derivadas.


### buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

Crea un intento desde la acción determinada. Debe implementarse mediante clases derivadas.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _createAttemptResult_|La acción de creación de intentos para el intento especificado.|
