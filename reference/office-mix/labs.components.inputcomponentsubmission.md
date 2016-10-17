
# <a name="labs.components.inputcomponentsubmission"></a>Labs.Components.InputComponentSubmission

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Representa un envío a un componente de entrada.

```
class InputComponentSubmission
```


## <a name="properties"></a>Propiedades


|Propiedad|Descripción|
|:-----|:-----|
| `public var answer: Components.InputComponentAnswer`|La respuesta ([Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)) asociada al envío.|
| `public var result: Components.InputComponentResult`|El resultado ([Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)) del envío.|
| `public var time: number`|La hora en la que se recibió el envío.|

## <a name="methods"></a>Métodos




### <a name="constructor"></a>constructor

 `function constructor(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, time: number)`

Crea una nueva instancia de la clase **InputComponentSubmission**.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _answer_|La respuesta asociada al envío.|
| _result_|El resultado del envío.|
| _time_|La hora en la que se recibió el envío.|
