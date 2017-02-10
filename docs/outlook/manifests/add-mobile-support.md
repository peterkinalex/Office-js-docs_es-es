# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Agregar compatibilidad a los comandos de complemento para Outlook Mobile

> **Nota**: Los comandos de complemento para Outlook Mobile solo se admiten en estos momentos en Outlook para iOS.

Usar comandos de complemento en Outlook Mobile permite que los usuarios tengan acceso a las mismas funciones (con algunas [limitaciones](#code-considerations)) que ya tienen en Outlook para Windows, Outlook para Mac y Outlook en la web. Para agregar compatibilidad para Outlook Mobile se necesita actualizar el manifiesto del complemento y, posiblemente, cambiar el código de los escenarios móviles.

## <a name="updating-the-manifest"></a>Actualizar el manifiesto

El primer paso para habilitar los comandos de complemento en Outlook Mobile es definirlos en el manifiesto del complemento. El esquema **VersionOverrides** v1.1 define un nuevo factor de forma para dispositivos móviles, [MobileFormFactor](../../reference/manifest/mobileformfactor.md).

Este elemento contiene toda la información para cargar el complemento en clientes móviles. Esto le permite definir completamente diferentes elementos de interfaz de usuario y archivos de JavaScript para la experiencia móvil.

En el siguiente ejemplo se muestra un único botón del panel de tareas en un elemento **MobileFormFactor**.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Control xsi:type="MobileButton" id="TaskPane1Btn">
        <Label resid="residTaskPaneButton0Name" />
        <Icon xsi:type="bt:MobileIconList">
          <bt:Image size="25" scale="1" resid="tp0icon" />
          <bt:Image size="25" scale="2" resid="tp0icon" />
          <bt:Image size="25" scale="3" resid="tp0icon" />

          <bt:Image size="32" scale="1" resid="tp0icon" />
          <bt:Image size="32" scale="2" resid="tp0icon" />
          <bt:Image size="32" scale="3" resid="tp0icon" />

          <bt:Image size="48" scale="1" resid="tp0icon" />
          <bt:Image size="48" scale="2" resid="tp0icon" />
          <bt:Image size="48" scale="3" resid="tp0icon" />
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="residTaskpaneUrl" />
        </Action>
      </Control>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

Esto es muy similar a los elementos que aparecen en un elemento [DesktopFormFactor](../../reference/manifest/desktopformfactor.md), con algunas diferencias notables.

- El elemento [OfficeTab](../../reference/manifest/officetab.md) no se usa.
- El elemento [ExtensionPoint](../../reference/manifest/exensionpoint.md) solo debe tener un elemento secundario. Si el complemento solo agrega un botón, el elemento secundario debe ser un elemento [Control](../../reference/manifest/control.md). Si el complemento agrega más de un botón, el elemento secundario debe ser un elemento [Group](../../reference/manifest/group.md) que contenga varios elementos `Control`.
- No existe ningún tipo `Menu` equivalente para el elemento `Control`.
- El elemento [Supertip](../../reference/manifest/supertip.md) no se usa.
- Los tamaños de icono necesarios son diferentes. Los complementos móviles deben admitir mínimamente iconos de 25x25, 32x32 y 48x48 píxeles.

## <a name="code-considerations"></a>Consideraciones de código

Diseñar un complemento para dispositivos móviles presenta algunas consideraciones adicionales.

### <a name="use-rest-instead-of-exchange-web-services"></a>Usar REST en lugar de servicios Web Exchange

El método [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) no se admite en Outlook Mobile. Los complementos deben preferir obtener información de la API de Office.js siempre que sea posible. Si los complementos necesitan información que no expone la API de Office.js, entonces deben usar las [API de REST de Outlook](https://dev.outlook.com/restapi/reference) para tener acceso al buzón del usuario. 

El conjunto de requisitos del buzón 1.5 presenta una versión nueva de [Office.context.mailbox.getCallbackTokenAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#getCallbackTokenAsync) que puede solicitar un token de acceso compatible con las API de REST, y una nueva propiedad [Office.context.mailbox.restUrl](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#restUrl) que puede usarse para buscar el punto de conexión de la API de REST para el usuario.

### <a name="pinch-zoom"></a>Gesto de acercamiento

De manera predeterminada, los usuarios pueden usar el "gesto de acercamiento" para ampliar los paneles de tareas. Si esto no tiene sentido en su escenario, asegúrese de deshabilitar el gesto de acercamiento en su HTML.

### <a name="closing-taskpanes"></a>Cerrar paneles de tareas

En Outlook Mobile, los paneles de tareas ocupan toda la pantalla y, de manera predeterminada, necesitan que el usuario los cierre para volver al mensaje. Considere la posibilidad de usar el método [Office.context.ui.closeContainer](https://dev.outlook.com/reference/add-ins/1.5/Office.context.ui.html#closeContainer) para cerrar el panel de tareas cuando el escenario esté completo.

### <a name="compose-mode-and-appointments"></a>Modo de redacción y citas

Actualmente, los complementos en Outlook Mobile solo admiten su activación al leer mensajes. Los complementos no están activados al redactar mensajes o al ver o redactar citas.

### <a name="unsupported-apis"></a>API no compatibles

Las siguientes API no son compatibles con Outlook Mobile.

  - [Office.context.officeTheme](../../reference/outlook/Office.context.md)
  - [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToEwsId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToRestId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.resources](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyAllForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntities](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatches](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)