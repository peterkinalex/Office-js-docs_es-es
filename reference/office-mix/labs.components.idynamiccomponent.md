
# Labs.Components.IDynamicComponent

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Permite la interacción con un componente dinámico.

```
interface IDynamicComponent extends Labs.Core.IComponent
```


## Propiedades


|Nombre|Descripción|
|:-----|:-----|
| `generatedComponentTypes: string[]`|Una matriz que contiene los tipos de componentes que este componente dinámico puede generar.|
| `maxComponents: number`|El número máximo de componentes que se generarán mediante este componente dinámico. O **Labs.Components.Infinite** si no existe ningún límite.|
