
# Labs.Components.ChoiceComponentAttempt

 _**Hace referencia a:** aplicaciones para Office | Complementos de Office | Office Mix | PowerPoint_

Representa un intento en un componente de elección.

```
class ChoiceComponentAttempt extends Components.ComponentAttempt
```


## Métodos




### constructor

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Crea una nueva instancia de la clase **ChoiceComponentAttempt**.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _labs_|La instancia [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) que se debe usar con el intento.|
| _attemptId_|El identificador asociado al intento.|
| _values_|Los valores asociados al intento.|

### timeout

 `public function timeout(callback: Labs.Core.ILabCallback<void>): void`

Indica que el laboratorio ha agotado el tiempo de espera.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Funciones de devolución de llamada que se desencadenan una vez que el servidor ha recibido el mensaje de tiempo de espera.|

### getSubmissions

 `public function getSubmissions(): Components.ChoiceComponentSubmission[]`

Recupera todos los envíos que se habían enviado previamente para un determinado intento.


### submit

 `public function submit(answer: Components.ChoiceComponentAnswer, result: Components.ChoiceComponentResult, callback: Labs.Core.ILabCallback<Components.ChoiceComponentSubmission>): void`

Envía una respuesta nueva que el laboratorio ha calificado y que no usará el host para calcular una calificación.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _answer_|La respuesta para el intento.|
| _result_|El resultado del envío.|
| _callback_|Función de devolución de llamada que se desencadena una vez que se ha recibido el envío.|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Inicia el procesamiento de la acción [Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md).

