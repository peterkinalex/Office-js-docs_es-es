
# Labs.Components.InputComponentResult

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

El resultado de un envío de un componente de entrada.

```
class InputComponentResult
```


## Propiedades


|Propiedad|Descripción|
|:-----|:-----|
| `public var score: any`|La puntuación asociada al envío.|
| `public var complete: boolean`|Indica si el resultado enviado provocó la finalización del intento.  **True** si el intento se ha completado.|

## Métodos




### constructor

 `function constructor(score: any, complete: boolean)`

Crea una nueva instancia de la clase **InputComponentResult**.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _score_|La puntuación asociada al resultado.|
| _completo_|Valor booleano **true** si el resultado ha completado el intento.|
