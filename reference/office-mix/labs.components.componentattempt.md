
# <a name="labs.components.componentattempt"></a>Labs.Components.ComponentAttempt

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Clase base para los intentos en los componentes.

```
class ComponentAttempt
```


## <a name="properties"></a>Propiedades


|**Nombre**|**Descripción**|
|:-----|:-----|
| `public var _componentId: string`|Identificador del componente especificado.|
| `public var _id: string`|Identificador del laboratorio asociado.|
| `public var _labs: Labs.LabsInternal`|El objeto de laboratorio ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) que se usa para interactuar con el objeto [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) subyacente.|
| `public var _resumed: boolean`|**True** si el laboratorio ha reanudado el progreso de un intento determinado.|
| `public var _state: Labs.ProblemState`|Estado actual del intento tal y como se ha proporcionado por la enumeración [Labs.ProblemState](../../reference/office-mix/labs.problemstate.md).|
| `public var _values: { [type:string]: Labs.ValueHolder<any>[]}`|Valores asociados al intento, si los hubiera, como se incluyen en el objeto [Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md).|

## <a name="methods"></a>Métodos




### <a name="constructor"></a>constructor

 `(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Crea una instancia nueva de la clase ComponentAttempt y proporciona valores del parámetro de entrada.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _labs_|La instancia [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) que se debe usar con el intento.|
| _attemptId_|El identificador asociado al intento.|
| _values_|Matriz de valores ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)) asociada al intento.|

### <a name="isresumed"></a>isResumed

 `public function isResumed(): boolean`

Función booleana que indica si el laboratorio se ha reanudado.  **True** si el laboratorio se ha reanudado.

 **Parámetros**

Ninguno.


### <a name="resume"></a>resume

 `public function resume(callback: Labs.Core.ILabCallback<void>): void`

Indica si el laboratorio ha reanudado el progreso del intento determinado y si carga los datos existentes como parte de este proceso. Un intento debe reanudarse antes de que se pueda usar.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena una vez que se ha reanudado el intento.|

### <a name="getstate"></a>getState

 `public function getState(): Labs.ProblemState`

Recupera el estado del laboratorio.

 **Parámetros**

Ninguno.


### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

Ejecuta la acción asociada al intento.

 **Parámetros**

Ninguno.


### <a name="getvalues"></a>getValues

 `public function getValues(key: string): Labs.ValueHolder<any>[]`

Recupera los valores asociados al intento.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _key_|La clave asociada al valor en la asignación de valores.|
