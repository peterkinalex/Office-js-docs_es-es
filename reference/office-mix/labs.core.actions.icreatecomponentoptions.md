
# Labs.Core.Actions.ICreateComponentOptions

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Crea un nuevo componente.

```
interface ICreateComponentOptions extends Core.IActionOptions
```


## Propiedades


|||
|:-----|:-----|
| `componentId: string`|El componente que invoca la acción de creación de componentes.|
| `component: Core.IComponent`|El componente [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md) que se debe crear|
| `correlationId?: string`|Campo opcional para establecer la correlación de este componente con todas las instancias de un laboratorio. Permite que el host identifique diferentes intentos en el mismo componente.|
