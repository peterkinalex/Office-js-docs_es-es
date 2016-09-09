
# Labs.Components.DynamicComponentInstance

 _**Hace referencia a:** aplicaciones para Office | Complementos de Office | Office Mix | PowerPoint_

Representa una instancia de un componente dinámico.

```
class DynamicComponentInstance extends Labs.ComponentInstanceBase
```


## Propiedades


|Propiedad|Descripción|
|:-----|:-----|
| `public var component: Components.IDynamicComponentInstance`|La definición de la instancia del componente.|

## Métodos




### constructor

 `function constructor(component: Components.IDynamicComponentInstance)`

Crea una nueva instancia del componente dinámico que usa la definición [Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md).


### getComponents

 `public function getComponents(callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase[]>): void`

Recupera todos los componentes creados por este componente dinámico.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena una vez que se han recuperado todos los componentes.|

### createComponent

 `public function createComponent(component: Labs.Core.IComponent, callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase>): void`

Crea un nuevo componente mediante el componente dinámico como base de este.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _componente_|El componente ([Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)) desde el que se crea la instancia.|
| _callback_|Función de devolución de llamada que se desencadena una vez que el componente se crea.|

### close

 `public function close(callback: Labs.Core.ILabCallback<void>): void`

Indica que no habrá envíos adicionales asociados a esta instancia del componente.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se desencadena una vez que la instancia se cierra.|

### isClosed

 `public function isClosed(callback: Labs.Core.ILabCallback<boolean>): void`

Devuelve un valor si el componente dinámico se cierra. Devuelve **True** si está cerrado.

