
# Labs.Components.IInputComponent

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Permite interactuar con un componente de entrada.

```
interface IInputComponent extends Labs.Core.IComponent
```


## Propiedades


|Nombre|Descripción|
|:-----|:-----|
| `maxScore: number`|La puntuación máxima permitida para el componente de entrada.|
| `timeLimit: number`|Límite de tiempo para el problema de entrada.|
| `hasAnswer: boolean`|**True** si el componente tiene una respuesta.|
| `answer: any`|La respuesta al problema del componente, si la hubiera.|
| `secure: boolean`|**True** si el componente de entrada es seguro.|
