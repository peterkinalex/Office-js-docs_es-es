
# <a name="labs.components.inputcomponentresult"></a>Labs.Components.InputComponentResult

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

El resultado de un envío de un componente de entrada.

```
class InputComponentResult
```


## <a name="properties"></a>Propiedades


|Propiedad|Descripción|
|:-----|:-----|
| `public var score: any`|La puntuación asociada al envío.|
| `public var complete: boolean`|Indica si el resultado enviado provocó la finalización del intento.  **True** si el intento se ha completado.|

## <a name="methods"></a>Métodos




### <a name="constructor"></a>constructor

 `function constructor(score: any, complete: boolean)`

Crea una nueva instancia de la clase **InputComponentResult**.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _score_|La puntuación asociada al resultado.|
| _complete_|Valor booleano **true** si el resultado ha completado el intento.|
