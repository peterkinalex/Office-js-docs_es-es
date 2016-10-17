
# <a name="labs.components.choicecomponentresult"></a>Labs.Components.ChoiceComponentResult

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

El resultado de un envío de componente de elección.

```
class ChoiceComponentResult
```


## <a name="properties"></a>Propiedades


|Propiedad|Descripción|
|:-----|:-----|
| `public var score: any`|La puntuación asociada al envío.|
| `public var complete: boolean`|Si el resultado ha completado o no el intento.  **True** si el resultado ha completado el intento.|

## <a name="methods"></a>Métodos




### <a name="constructor"></a>constructor

 `function constructor(score: any, complete: boolean)`

Crea una nueva instancia de la clase **ChoiceComponentResult**.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _score_|La puntuación del resultado.|
| _complete_|Indica si el resultado ha completado el intento.|
