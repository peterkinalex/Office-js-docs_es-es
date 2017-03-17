# <a name="implement-a-pinnable-taskpane-in-outlook"></a>Implementar un panel de tareas anclable en Outlook

La forma de experiencia de usuario [panel de tareas](../add-in-commands-for-outlook.md#launching-a-task-pane) de los comandos de complemento abre un panel de tareas vertical en la parte derecha de los mensajes o citas abiertos. Esto permite que el complemento proporcione interfaz de usuario para realizar interacciones más detalladas (por ejemplo, rellenar varios campos). Este panel de tareas puede mostrarse en el panel de lectura al ver una lista de mensajes, lo que permite procesar rápido los mensajes.

Sin embargo, de forma predeterminada, si un usuario tiene el panel de tareas de un complemento abierto para un mensaje en el panel de lectura y, a continuación, selecciona un mensaje nuevo, el panel de tareas se cerrará automáticamente. Es posible que, en el caso de un complemento que se use mucho, el usuario prefiera mantener el panel abierto para eliminar la necesidad de volver a activar el complemento en cada mensaje. Con los paneles de tareas anclables, el complemento puede ofrecer al usuario esa opción.

> **Nota**: Los paneles de tareas anclables solo se admiten en Outlook 2016 para Windows (compilación 7668.2000 o posterior para los usuarios del Canal actual u Office Insider, compilación 7900.xxxx o posterior para los usuarios del Canal Diferido).

## <a name="support-taskpane-pinning"></a>Admitir el anclado de paneles de tareas

El primer paso es agregar la compatibilidad con el anclado (debe realizarse en el [manifiesto](./manifests.md) del complemento). Esta acción se realiza agregando el elemento [SupportsPinning](../../../reference/manifest/action.md#supportspinning) al elemento `Action` que describe el botón del panel de tareas.

El elemento `SupportsPinning` se define en el esquema de la versión 1.1 de VersionOverrides, así que deberá incluir un elemento [VersionOverrides](../../../reference/manifest/versionoverrides.md) tanto para la versión 1.0 como para la 1.1.

> **Nota**: Si tiene pensado [publicar](../../publish/publish.md) su complemento de Outlook en la Tienda Office, cuando utilice el elemento **SupportsPinning**, para poder superar la [validación de la Tienda Office](https://msdn.microsoft.com/en-us/library/jj220035.aspx), el contenido del complemento no puede ser estático y debe mostrar claramente los datos relacionados con el mensaje que está abierto o seleccionado en el buzón de correo.

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

Para obtener un ejemplo completo, consulte el control `msgReadOpenPaneButton` del [manifiesto de ejemplo command-demo](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml).

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>Control de las actualizaciones de la interfaz de usuario en función del mensaje seleccionado en ese momento

Para actualizar la interfaz de usuario o las variables internas de su panel de tareas en función del elemento actual, deberá registrar un controlador de eventos para que se le notifique el cambio.

### <a name="implement-the-event-handler"></a>Implementar el controlador de eventos

El controlador de eventos debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` de este objeto se establecerá en `Office.EventType.ItemChanged`. Cuando se llama al evento, el objeto `Office.context.mailbox.item` ya se ha actualizado para reflejar el elemento que hay seleccionado actualmente.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

### <a name="register-the-event-handler"></a>Registrar el controlador de eventos

Use el método [Office.context.mailbox.addHandlerAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#addHandlerAsync) para registrar el controlador de eventos del evento `Office.EventType.ItemChanged`. Esta acción debe realizarse en la función `Office.initialize` de su panel de tareas.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="additional-resources"></a>Recursos adicionales

Para obtener un complemento de ejemplo que implemente un panel de tareas anclable, consulte [command-demo](https://github.com/jasonjoh/command-demo) en GitHub.