
# <a name="labs.components.iinputcomponent"></a>Labs.Components.IInputComponent

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Permite la interacción con un componente de entrada.

```
interface IInputComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>Propiedades


|Nombre|Descripción|
|:-----|:-----|
| `maxScore: number`|La puntuación máxima permitida para el componente de entrada.|
| `timeLimit: number`|Límite de tiempo para el problema de entrada.|
| `hasAnswer: boolean`|**True** si el componente tiene una respuesta.|
| `answer: any`|La respuesta al problema del componente, si la hubiera.|
| `secure: boolean`|**True** si el componente de entrada es seguro.|
