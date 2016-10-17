
# <a name="document.settings-property"></a>Propiedad Document.settings
Obtiene un objeto que representa la configuración personalizada que se ha guardado del complemento de contenido o del panel de tareas para el documento actual.

|||
|:-----------------|:--------------------------------|
| Hosts:           | Access, Excel, PowerPoint y Word |
| Modificado por última vez en: | 1.1                             |

```js
var _settings = Office.context.document.settings;
```

## <a name="return-value"></a>Valor devuelto

Un objeto [Settings](./settings.md).

## <a name="support-details"></a>Detalles de compatibilidad

Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).

**Hosts compatibles, por plataforma**

|             | Office para escritorio de Windows | Office Online (en el explorador) | Office para iPad |
|:------------|:---------------------------|:---------------------------|:----------------|
| Access      |                            | v                          |                 |
| Excel       | v                          | v                          | v               |
| PowerPoint  | v                          | v                          | v               |
| Word        | v                          | v                          | v               |

|||
|:--------------------------|:-----|
| Nivel de permisos mínimo  | [Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
| Tipos de complementos:             | Contenido, panel de tareas
| Biblioteca:                  | Office.js
| Espacio de nombres:                | Office

## <a name="support-history"></a>Historial de compatibilidad

| Versión | Cambios |
|:--------|:--------|
| 1.1     |Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.
| 1.1     |Se ha agregado compatibilidad con complementos de contenido para Access.
| 1.0     |Agregado
