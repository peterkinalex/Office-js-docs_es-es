# Elemento FormFactor

Especifica la configuración para un complemento para un factor de forma determinado. Por ejemplo, definir un `Host` con los tipos `MailHost` y `DesktopFormFactor` se aplicará a Outlook para escritorio pero _no_ a Outlook Web App o Outlook.com. Contiene toda la información de complemento para dicho factor de forma excepto para el nodo **Resources**.

Cada definición FormFactor contiene el elemento **FunctionFile** y uno o más elementos **ExtensionPoint**. Para obtener más información, consulte [Elemento FunctionFile](./functionfile.md) y [Elemento ExtensionPoint](./extensionpoint.md). 

Se admiten los siguientes FormFactors:

- `DesktopFormFactor` (Office para clientes Windows o Mac)

## Elementos secundarios

| Elemento                               | Obligatorio | Descripción  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | Sí      | Define dónde expone su funcionalidad un complemento. |
| [FunctionFile](./functionfile.md)     | Sí      | Una dirección URL de un archivo que contiene funciones de JavaScript.|
| [GetStarted](./getstarted.md)         | No       | Define la llamada que aparece cuando se instala el complemento en hosts de Word, Excel o PowerPoint. |

## Ejemplo de FormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
