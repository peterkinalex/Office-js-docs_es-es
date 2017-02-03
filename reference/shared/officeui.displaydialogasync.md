# <a name="uidisplaydialogasync-method"></a>Método UI.displayDialogAsync

Muestra un cuadro de diálogo en un host de Office. 

## <a name="requirements"></a>Requisitos

|Host|Incorporación|Modificado por última vez en|
|:---------------|:--------|:----------|
|Word, Excel, PowerPoint|1.1|1.1|
|Outlook|Buzón 1.4|Buzón 1.4|

Este método está disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md) DialogApi para los complementos de Word, Excel o PowerPoint, así como en el conjunto de requisitos de Buzón 1.4 para Outlook. Para especificar el conjunto de requisitos de DialogAPI, use lo siguiente en su manifiesto.

```xml
<Requirements> 
  <Sets DefaultMinVersion="1.1"> 
    <Set Name="DialogApi"/> 
  </Sets> 
</Requirements> 
```

Para especificar el conjunto de requisitos de Buzón 1.4, use lo siguiente en su manifiesto.

```xml
<Requirements> 
  <Sets DefaultMinVersion="1.4"> 
    <Set Name="Mailbox"/> 
  </Sets> 
</Requirements> 
```

Para detectar esta API en tiempo de ejecución en un complemento de Word, Excel o PowerPoint, use el siguiente código.

```js
if (Office.context.requirements.isSetSupported('DialogApi', 1.1)) {  
  // Use Office UI methods; 
} else { 
  // Alternate path 
} 
```

Para detectar esta API en tiempo de ejecución en un complemento de Outlook, use el siguiente código.

```js
if (Office.context.requirements.isSetSupported('Mailbox', 1.4)) {  
  // Use Office UI methods; 
} else { 
  // Alternate path 
} 
```

Como alternativa, puede comprobar si el método `displayDialogAsync` está sin definir antes de usarlo.

```js
if (Office.context.ui.displayDialogAsync !== undefined) {
  // Use Office UI methods
}
```

### <a name="supported-platforms"></a>Plataformas compatibles
Para obtener información acerca de las plataformas compatibles, consulte [Conjuntos de requisitos de la API de cuadros de diálogo](../requirement-sets/dialog-api-requirement-sets.md).

## <a name="syntax"></a>Sintaxis

```js
Office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##<a name="examples"></a>Ejemplos

Para ver un ejemplo sencillo que usa el método **displayDialogAsync**, consulte [Office Add-in Dialog API example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/) (Ejemplo del complemento de Office Dialog API) en GitHub.

Para ver un ejemplo que muestre escenarios de autenticación, consulte:

- [Complemento de PowerPoint en gráfico de inserción de Microsoft Graph ASP.Net](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Complemento de Office Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Complemento de Excel ASP.NET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Ejemplo de autenticación de servidor de complementos de Office para ASP.net MVC](https://github.com/dougperkes/Office-Add-in-AspNetMvc-ServerAuth/tree/Office2016DisplayDialog)
- [Autenticación de cliente de Office 365 de complementos de Office para AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)


 
## <a name="parameters"></a>Parámetros

| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|startAddress|cadena|Acepta la dirección URL HTTPS(TLS) inicial que se abre en el cuadro de diálogo. <ul><li>La página inicial debe estar en el mismo dominio que la página primaria. Después de cargar la página inicial, puede ir a otros dominios.</li><li>Cualquier página que llame a [office.context.ui.messageParent](officeui.messageparent.md) también debe estar en el mismo dominio que la página principal.</li></ul>|
|options|object|Opcional. Acepta un objeto de opciones para definir los comportamientos de los cuadros de diálogo.|
|callback|objeto|Acepta un método de devolución de llamada para controlar el intento de creación de cuadro de diálogo.|
    
### <a name="configuration-options"></a>Opciones de configuración
Las opciones de configuración siguientes están disponibles para un cuadro de diálogo.


| Propiedad     | Tipo   |Descripción|
|:---------------|:--------|:----------|
|**width**|int|Opcional. Define el ancho del cuadro de diálogo como porcentaje de la pantalla actual. El valor predeterminado es 80 %. La resolución mínima es de 250 píxeles.|
|**height**|int|Opcional. Define la altura del cuadro de diálogo como porcentaje de la pantalla actual. El valor predeterminado es 80 %. La resolución mínima es de 150 píxeles.|
|**displayInIframe**|bool|Opcional. Determina si se debe mostrar el cuadro de diálogo dentro de un IFrame. **Esta configuración solo se aplica en clientes de Office Online**, se omite para los clientes de escritorio. Los valores posibles son:<ul><li>falso (predeterminado): se mostrará el cuadro de diálogo como una nueva ventana de explorador (elemento emergente). Se recomienda para las páginas de autenticación que no se pueden mostrar en un IFrame. </li><li>verdadero: se mostrará el cuadro de diálogo como una superposición flotante con un IFrame. Es la mejor opción para mejorar el rendimiento y la experiencia del usuario.</li>|


## <a name="callback-value"></a>Valor de devolución de llamada
Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **displayDialogAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Acceso al objeto [Dialog](../../reference/shared/officeui.dialog.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Acceda al valor u objeto definidos por el usuario si ha pasado uno como parámetro _asyncContext_.|

### <a name="errors-from-displaydialogasync"></a>Errores de displayDialogAsync

Además de los errores del sistema y de la plataforma en general, los siguientes errores son específicos para llamar a **displayDialogAsync**.

|**Número de código**|**Significado**|
|:-----|:-----|
|12004|El dominio de la dirección URL pasado a `displayDialogAsync` no es de confianza. El dominio debe estar en el mismo dominio que la página de host (incluido el número de protocolo y de puerto), o debe registrarse en la sección `<AppDomains>` del manifiesto del complemento.|
|12005|La dirección URL pasada a `displayDialogAsync` utiliza el protocolo HTTP. Se necesita HTTPS. (En algunas versiones de Office, el mensaje de error devuelto con 12005 es el mismo devuelto para 12004).|
|12007|Ya hay un cuadro de diálogo abierto en el panel de tareas. Un complemento de panel de tareas solo puede tener abierto un cuadro de diálogo al mismo tiempo.|



## <a name="design-considerations"></a>Consideraciones sobre diseño
Se aplican las siguientes consideraciones de diseño a cuadros de diálogo:

- Un complemento de Office puede tener solo un cuadro de diálogo abierto en cualquier momento.
- El usuario puede mover y cambiar de tamaño cada cuadro de diálogo.
- Cada cuadro de diálogo se centra en la pantalla cuando se abre.
- Los cuadros de diálogo aparecen encima de la aplicación host y en el orden en que se crearon.

Use un cuadro de diálogo para:

- Mostrar páginas de autenticación para recopilar las credenciales de usuario.
- Mostrar una pantalla de progreso/error/entrada de un comando ShowTaspane o ExecuteAction.
- Aumentar temporalmente el área de superficie que un usuario tiene disponible para completar una tarea.

No use un cuadro de diálogo para interactuar con un documento. En su lugar, use un panel de tareas. 

Para obtener un modelo de diseño que pueda usar para crear un cuadro de diálogo, consulte [Client Dialog](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md) en el repositorio de modelos de diseño de experiencia de usuario de complementos de Office en GitHub.
