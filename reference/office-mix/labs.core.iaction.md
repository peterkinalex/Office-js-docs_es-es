
# <a name="labs.core.iaction"></a>Labs.Core.IAction

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Representa una acción de laboratorio, que es una interacción que un usuario tiene con un laboratorio especificado.

```
interface IAction
```


## <a name="properties"></a>Properties


|||
|:-----|:-----|
| `type: string`|El tipo de acción realizada por el usuario.|
| `options: Core.IActionOptions`|Las opciones [Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md) enviadas con la acción realizada por el usuario.|
| `result: Core.IActionResult`|El resultado [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) de la acción.|
| `time: number`|La hora a la que se completó la acción, representada en milisegundos transcurridos desde el 1 de enero de 1970 00:00:00 UTC.|
