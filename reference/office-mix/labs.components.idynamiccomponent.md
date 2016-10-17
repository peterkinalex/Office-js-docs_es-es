
# <a name="labs.components.idynamiccomponent"></a>Labs.Components.IDynamicComponent

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Permite la interacción con un componente dinámico.

```
interface IDynamicComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>Propiedades


|Nombre|Descripción|
|:-----|:-----|
| `generatedComponentTypes: string[]`|Una matriz que contiene los tipos de componentes que este componente dinámico puede generar.|
| `maxComponents: number`|El número máximo de componentes que se generarán mediante este componente dinámico. O **Labs.Components.Infinite** si no existe ningún límite.|
