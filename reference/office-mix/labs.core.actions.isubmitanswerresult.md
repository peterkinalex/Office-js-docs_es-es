
# <a name="labs.core.actions.isubmitanswerresult"></a>Labs.Core.Actions.ISubmitAnswerResult

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

El resultado de enviar una respuesta para un intento.

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## <a name="properties"></a>Properties


|||
|:-----|:-----|
| `submissionId: string`|Un identificador asociado al envío. Proporcionado por el servidor.|
| `complete: boolean`|Devuelve **True** si el intento se completa debido al envío actual.|
| `score: any`|Información de puntuación asociada al envío.|
