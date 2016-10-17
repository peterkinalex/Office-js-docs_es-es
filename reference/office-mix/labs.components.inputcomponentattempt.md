
# <a name="labs.components.inputcomponentattempt"></a>Labs.Components.InputComponentAttempt

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Representa un intento de interactuar con un componente de entrada.

```
class InputComponentAttempt extends Components.ComponentAttempt
```


## <a name="methods"></a>Métodos




### <a name="constructor"></a>constructor

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Crea una nueva instancia de la clase **InputComponentAttempt**.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _labs_|Los laboratorios ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) asociados al intento.|
| _componentID_|Identificador del componente asociado al intento.|
| _attemptId_|Identificador del intento específico.|
| _values_|Una matriz que contiene las instancias del valor ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)).|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

Recorre en iteración las acciones recuperadas para el intento especificado y rellena el estado del laboratorio.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _action_|Acción asociada al estado del laboratorio.|

### <a name="getsubmissions"></a>getSubmissions

 `public function getSubmissions(): Components.InputComponentSubmission[]`

Recupera todos los envíos que se han enviado previamente para el intento especificado.


### <a name="submit"></a>submit

 `public function submit(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, callback: Labs.Core.ILabCallback<Components.InputComponentSubmission>): void`

Envía una respuesta nueva que el laboratorio ha calificado y que no usará el host para calcular una calificación.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _answer_|La respuesta asociada al intento.|
| _result_|El resultado asociado al envío.|
| _callback_|Función de devolución de llamada que se desencadena una vez que se ha recibido el envío.|
