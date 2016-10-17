
# <a name="labs.core.icomponentinstance"></a>Labs.Core.IComponentInstance

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Clase base de las instancias de los componentes del laboratorio.

```
interface IComponentInstance extends Core.ILabObject, Core.IUserData
```


## <a name="properties"></a>Propiedades


|||
|:-----|:-----|
| `componentId: string`|El identificador del componente al que está asociada esta instancia.|
| `name: string`|Nombre del componente.|
| `values: {[type:string]: Core.IValueInstance[]}`|La asignación de propiedades de valor asociada al componente.|

## <a name="remarks"></a>Observaciones

Una instancia de componente es una creación de instancia de un componente para un usuario. Contiene una vista traducida del componente para una ejecución particular del laboratorio. Esta vista puede excluir información oculta (respuestas y sugerencias entre otros) y también contiene identificadores para identificar las diversas instancias.

