
# Labs.Core.IComponent

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Clase base para representar componentes de un laboratorio.

```
interface IComponent extends Core.ILabObject, Core.IUserData
```


## Propiedades


|||
|:-----|:-----|
| `name: string`|Nombre del componente.|
| `values: {[type:string]: Core.IValue[]}`|La asignación de propiedades de valor asociada al componente.|
