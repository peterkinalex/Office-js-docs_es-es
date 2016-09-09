
# Labs.Components.IChoiceComponentInstance

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Una instancia de un componente de elección.

```
interface IChoiceComponentInstance extends Labs.Core.IComponentInstance
```


## Propiedades


|Nombre|Descripción|
|:-----|:-----|
| `choices: Components.IChoice[]`|Una matriz que representa el listado de elecciones asociadas al problema.|
| `timeLimit: number`|Límite de tiempo para completar el problema.|
| `maxAttempts: number`|Número máximo de intentos permitidos para el problema.|
| `maxScore: number`|La puntuación máxima del problema.|
| `hasAnswer: boolean`|**True** si el problema tiene una respuesta.|
| `answer: any`|La respuesta del problema. Una matriz si se admiten varias respuestas o un identificador único si solo se admite una respuesta.|
| `secure: boolean`|Si el cuestionario es seguro o no, lo que significa que el usuario retiene los campos seguros.|
