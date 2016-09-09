
# Labs.Core.Actions.ISubmitAnswerResult

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

El resultado de enviar una respuesta para un intento.

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## Properties


|||
|:-----|:-----|
| `submissionId: string`|Un identificador asociado al envío. Proporcionado por el servidor.|
| `complete: boolean`|Devuelve **True** si el intento se completa debido al envío actual.|
| `score: any`|Información de puntuación asociada al envío.|
