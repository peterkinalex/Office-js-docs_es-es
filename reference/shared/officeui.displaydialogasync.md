# Método UI.displayDialogAsync

Muestra un cuadro de diálogo en un host de Office. 

## Requisitos

|Host|Incorporación|Modificado por última vez en|
|:---------------|:--------|:----------|
|Word, Excel, PowerPoint|1.1|1.1|
|Outlook|Buzón 1.4|Buzón 1.4|

Este método está disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md) de DialogAPI. Para especificar el conjunto de requisitos de DialogAPI, use lo siguiente en su manifiesto.

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="DialogAPI"/> 
   </Sets> 
 </Requirements> 

```

Para detectar esta API en tiempo de ejecución, use el siguiente código.

```js
 if (Office.context.requirements.isSetSupported('DialogAPI', 1.1)) 
    {  
         // Use Office UI methods; 
    } 
 else 
     { 
         // Alternate path 
     } 
```



### Plataformas compatibles
El conjunto de requisitos de DialogAPI actualmente es compatible con las siguientes plataformas:

  - Office 2016 para escritorio de Windows (versión 16.0.6741.0000 o posteriores)
  - Office para IPad (versión 1.22 o posteriores)
  - Office para Mac (versión 15.20 o posteriores) 

Estarán disponibles en más plataformas próximamente. 

## Sintaxis

```js
office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##Ejemplos

Para ver un ejemplo simple que usa el método **displayDialogAsync**, consulte [Office Add-in Dialog API example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/) (Ejemplo del complemento de Office Dialog API) en GitHub.

Para ver un ejemplo que muestra un escenario de autenticación, consulte el ejemplo [Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth) (Complemento de Office para autenticación de cliente de Office 365 para AngularJS) en GitHub.

 
## Parámetros

| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|startAddress|cadena|Acepta la dirección URL HTTPS(TLS) inicial que se abre en el cuadro de diálogo. <ul><li>La página inicial debe estar en el mismo dominio que la página primaria. Después de cargar la página inicial, puede ir a otros dominios.</li><li>Cualquier página que llame a [office.context.ui.messageParent](officeui.messageparent.md) también debe estar en el mismo dominio que la página principal.</li></ul>|
|options|object|Opcional. Acepta un objeto de opciones para definir los comportamientos de los cuadros de diálogo.|
|callback|objeto|Acepta un método de devolución de llamada para controlar el intento de creación de cuadro de diálogo.|
    
### Opciones de configuración
Las opciones de configuración siguientes están disponibles para un cuadro de diálogo.


| Propiedad     | Tipo   |Descripción|
|:---------------|:--------|:----------|
|**width**|objeto|Opcional. Define el ancho del cuadro de diálogo como porcentaje de la pantalla actual. El valor predeterminado es 80 %. La resolución mínima es de 250 píxeles.|
|**height**|objeto|Opcional. Define la altura del cuadro de diálogo como porcentaje de la pantalla actual. El valor predeterminado es 80 %. La resolución mínima es de 150 píxeles.|
|**displayInIFrame**|object|Opcional. Determina si se debe mostrar el cuadro de diálogo dentro de un IFrame en clientes de Office Online. Esta configuración se omite por los clientes de escritorio. Los valores posibles son:<ul><li>False (predeterminado): se mostrará el cuadro de diálogo como una nueva ventana de explorador (elemento emergente). Se recomienda para las páginas de autenticación que no se pueden mostrar en un IFrame. </li><li>True: se mostrará el cuadro de diálogo como una superposición flotante con un IFrame. Es adecuado para mejorar el rendimiento y la experiencia del usuario.</li>|


## Valor de devolución de llamada
Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **displayDialogAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Acceso al objeto [Dialog](../../reference/shared/officeui.dialog.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Acceda al valor o al objeto definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|


## Consideraciones sobre diseño
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
