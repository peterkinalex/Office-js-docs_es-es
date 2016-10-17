
# <a name="officetheme.controlbackgroundcolor-property"></a>Propiedad officeTheme.controlBackgroundColor
Obtiene el color de fondo del control del tema de Office.

 **Importante:** Actualmente, esta API solo funciona en Excel, Outlook, PowerPoint y Word en [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) para el escritorio de Windows.



|||
|:-----|:-----|
|**Hosts:**|Excel, Outlook, PowerPoint y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|No en un conjunto|
|**Agregado en**|1.3|

```
var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
```


## <a name="return-value"></a>Valor devuelto

Un triplo de color hexadecimal.


## <a name="remarks"></a>Observaciones

Los colores devueltos se corresponden con los valores del tema de Office seleccionado por el usuario en **Archivo**  >  **Cuenta de Office**  >  interfaz de usuario de **Tema de Office**, que se aplica a todas las aplicaciones host de Office.


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|**OWA para dispositivos**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|v||||
|**Outlook**|v||||
|**PowerPoint**|v||||
|**Word**|v||||

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas y Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.3|Agregado|
