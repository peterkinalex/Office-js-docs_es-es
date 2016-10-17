
# <a name="labs.core.icomponent"></a>Labs.Core.IComponent

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Clase base para representar componentes de un laboratorio.

```
interface IComponent extends Core.ILabObject, Core.IUserData
```


## <a name="properties"></a>Propiedades


|||
|:-----|:-----|
| `name: string`|Nombre del componente.|
| `values: {[type:string]: Core.IValue[]}`|La asignación de propiedades de valor asociada al componente.|
