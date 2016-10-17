
# <a name="context.officetheme-property"></a>Propiedad Context.officeTheme
Proporciona acceso a las propiedades de los colores del tema de Office.

 **Importante:** Actualmente, esta API solo funciona en Excel, Outlook, PowerPoint y Word en [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) para el escritorio de Windows.


|||
|:-----|:-----|
|**Hosts:**|Excel, Outlook, PowerPoint y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|No en un conjunto|
|**Agregado en**|1.3|



```js
Office.context.officeTheme
```


## <a name="members"></a>Miembros


**Propiedades**

|||
|:-----|:-----|
|Nombre|Descripción|
|[bodyBackgroundColor ](../../reference/shared/office.context.bodybackgroundcolor.md)|Obtiene el color de fondo del cuerpo del tema de Office.|
|[bodyForegroundColor](../../reference/shared/office.context.bodyforegroundcolor.md)|Obtiene el color de primer plano del cuerpo del tema de Office.|
|[controlBackgroundColor](../../reference/shared/office.context.controlbackgroundcolor.md)|Obtiene el color de fondo del control del tema de Office.|
|[controlForegroundColor](../../reference/shared/office.context.controlforegroundcolor.md)|Obtiene el color de primer plano del control del tema de Office.|

## <a name="remarks"></a>Observaciones

El uso de los colores del tema de Office le permite coordinar la combinación de colores del complemento con el tema actual de Office seleccionado por el usuario mediante **Archivo**  >  **Cuenta de Office**  >  interfaz de usuario **Tema de Office**, que se aplica a todas las aplicaciones host de Office. El uso de colores del tema de Office es idóneo para Outlook y para los complementos de panel de tareas.


## <a name="example"></a>Ejemplo


```js
function applyOfficeTheme(){
    // Get office theme colors.
    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply body background color to a CSS class.
    $('.body').css('background-color', bodyBackgroundColor);
}
```


## <a name="support-details"></a>Detalles de compatibilidad



|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas y Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad


|**Versión**|**Cambios**|
|:-----|:-----|
|1.3|Agregado|
