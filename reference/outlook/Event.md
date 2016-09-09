

# Evento

El objeto `event` se pasa como parámetro a las funciones del complemento que invocan los botones de comando directos. El objeto permite que el complemento identifique en qué botón se ha hecho clic y que indique el host que ha completado su procesamiento.

Por ejemplo, considere un botón definido en un manifiesto del complemento de la forma siguiente:

```
<Control xsi:type="Button" id="eventTestButton">
  <Label resid="eventButtonLabel" />
  <Tooltip resid="eventButtonTooltip" />
  <Supertip>
    <Title resid="eventSuperTipTitle" />
    <Description resid="eventSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>testEventObject</FunctionName>
  </Action>
</Control>
```

El botón tiene un atributo `id` establecido en `eventTestButton` e invocará la función `testEventObject` que se define en el complemento. Esa función tiene este aspecto:

```
function testEventObject(event) {
  // The event object implements the Event interface

  // This value will be "eventTestButton"
  var buttonId = event.source.id;

  // Signal to the host app that processing is complete.
  event.completed();
}
```

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

### Miembros

####  source :Object

Obtiene el identificador del botón de comando del complemento que ha invocado el método.

La propiedad `source` devuelve un objeto con las siguientes propiedades.

| Propiedad | Descripción |
| --- | --- |
| `id` | El valor del atributo `id` del elemento `Control` que define el botón de comando del complemento en el manifiesto del complemento. |

Este valor puede usarse cuando más de un botón invoca la misma función, pero necesita realizar acciones diferentes basándose en el botón en el que se ha hecho clic.

##### Tipo:

*   Objeto

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

##### Ejemplo

```
// Function is used by two buttons:
// button1 and button2
function multiButton (event) {
  // Check which button was clicked
  var buttonId = event.source.id;

  if (buttonId === 'button1') {
    doButton1Action();
  else {
    doButton2Action();
  }

  event.completed();
}
```

### Métodos

####  completed()

Indica que el complemento ha completado el procesamiento que se desencadenó mediante un botón de comando del complemento.

Este método debe llamarse al final de una función que se ha invocado mediante un comando de complemento definido con un elemento `Action` con un atributo `xsi:type` establecido en `ExecuteFunction`. Al llamar a este método se indica al cliente de host que la función está completa y que puede limpiar cualquier estado que esté implicado al invocar la función. Por ejemplo, si el usuario cierra Outlook antes de que se llame a este método, Outlook le advertirá de que una función continúa ejecutándose.

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

##### Ejemplo

```
function processItem (event) {
  // Do some processing

  event.completed();
}
```