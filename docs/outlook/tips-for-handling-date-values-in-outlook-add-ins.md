
# Sugerencias para el tratamiento de los valores de fecha en los complementos de Outlook

La API de JavaScript para Office usa el objeto de JavaScript [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp) para la mayoría de las operaciones de almacenamiento y recuperación de fechas y horas. El objeto **Date** proporciona métodos como [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp) y [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp), que devuelven el valor de fecha u hora solicitado según la hora UTC (hora universal coordinada).<br/><br/>
El objeto **Date** también proporciona otros métodos, como [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp) y [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp), que devuelven la fecha u hora solicitada según la "hora local".<br/><br/>El concepto de "hora local" se determina en gran medida por el explorador y el sistema operativo del equipo cliente. Por ejemplo, en la mayoría de los exploradores que se ejecutan en un equipo cliente basado en Windows, una llamada de JavaScript a **getDate** devuelve una fecha según la zona horaria configurada en Windows en el equipo cliente.<br/><br/>
En el ejemplo siguiente se crea un objeto **Date**<code>myLocalDate</code> en la hora local, que llama a **toUTCString** para convertir esa fecha en una cadena de fecha en UTC.




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Aunque puede usar el objeto de JavaScript  **Date** para obtener un valor de fecha u hora basado en UTC o en la zona horaria del equipo cliente, el objeto **Date** tiene una limitación: no proporciona métodos que devuelvan un valor de fecha u hora de ninguna otra zona horaria específica. Por ejemplo, si su equipo cliente está configurado con la Hora estándar del Este (EST), no hay ningún método **Date** que le permita obtener valores de hora en otras zonas horarias aparte de EST o UTC como, por ejemplo, la Hora estándar del Pacífico (PST).


## Funciones relacionadas con fechas para complementos de Outlook


La limitación de JavaScript mencionada tiene una implicación importante, cuando se usa la API de JavaScript para Office para controlar los valores de fecha u hora en los complementos de Outlook que se ejecutan en un cliente enriquecido de Outlook y en Outlook Web App o en OWA para dispositivos.


### Zonas horarias para clientes de Outlook

Para mayor claridad, definamos el concepto de zonas horarias.



|**Zona horaria**|**Descripción**|
|:-----|:-----|
|Zona horaria del equipo cliente|Esto se define en el sistema operativo del equipo cliente. La mayoría de los exploradores usan la zona horaria del equipo cliente para mostrar valores de fecha u hora del objeto **Date** de JavaScript.<br/><br/>Un cliente avanzado de Outlook usa esta zona horaria para mostrar valores de fecha u hora en la interfaz de usuario. <br/><br/>Por ejemplo, en un equipo cliente que ejecuta Windows, Outlook usa la zona horaria configurada en Windows como la zona horaria local. En Mac, si el usuario cambia la zona horaria en el equipo cliente, Outlook para Mac solicitará al usuario que actualice también la zona horaria en Outlook.|
|Zona horaria del Centro de administración de Exchange (EAC)|El usuario define este valor de zona horaria (así como el idioma preferido) al iniciar sesión por primera vez en Outlook Web App o en OWA para dispositivos. <br/><br/>Outlook Web App y OWA para dispositivos usan esta zona horaria para mostrar valores de fecha u hora en la interfaz de usuario.|
Puesto que un cliente enriquecido de Outlook usa la zona horaria del equipo cliente y la interfaz de usuario de Outlook Web App y OWA para dispositivos usa la zona horaria de EAC, la hora local del mismo complemento de correo instalado para el mismo buzón puede ser distinta cuando se ejecute en un cliente enriquecido de Outlook y en Outlook Web App o en OWA para dispositivos. Como desarrollador de complementos de Outlook, debe tomar y devolver los valores de fecha de manera apropiada de modo que esos valores sean siempre coherentes con la zona horaria esperada por el usuario en el cliente correspondiente.


### API relacionadas con fechas

Estas son las propiedades y los métodos de la API de JavaScript para Office que son compatibles con características relacionadas con fechas.reference/outlook/Office.context.mailbox.item.md



**Miembro de la API**|**Representación de zona horaria**|**Ejemplo en un cliente enriquecido de Outlook**|**Ejemplo en Outlook Web App o en OWA para dispositivos**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|En un cliente enriquecido de Outlook, esta propiedad devuelve la zona horaria del equipo cliente. En Outlook Web App y OWA para dispositivos, esta propiedad devuelve la zona horaria de EAC. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) y [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Cada una de estas propiedades devuelve un objeto **Date** de JavaScript. Este valor de **Date** es válido según UTC, como se muestra en el ejemplo siguiente: `myUTCDate` tiene el mismo valor en un cliente avanzado de Outlook, en Outlook Web App y en OWA para dispositivos.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Pero, al llamar a `myDate.getDate`, se devuelve un valor de fecha en la zona horaria del equipo cliente, que es coherente con la zona horaria usada para mostrar los valores de fecha y hora en la interfaz de clic de Outlook, pero puede ser distinta de la zona horaria EAC que Outlook Web App y OWA para dispositivos usan en su interfaz de usuario.|Si el elemento se creó a las 9:00 UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` devuelve 04:00 EST.<br/><br/>Si el elemento se modificó a las 11:00 UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` devuelve 06:00 EST.|Si la hora de creación del elemento es 9:00 UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` devuelve 04:00 EST.<br/><br/>Si el elemento se modificó a las 11:00 UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` devuelve 06:00 EST.<br/><br/>Tenga en cuenta que, si quiere mostrar la hora de creación o de modificación en la interfaz de usuario, primero tendría que convertir la hora a PST para que fuera coherente con el resto de la interfaz de usuario.
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|Los parámetros  _Start_ y _End_ requieren cada uno un objeto **Date** de JavaScript. Los argumentos deben ser correctos en términos de hora UTC, independientemente de cuál sea la zona horaria usada en la interfaz de usuario de un cliente enriquecido de Outlook, Outlook Web App o OWA para dispositivos.|Si las horas de inicio y de finalización del formulario de cita son 9:00 UTC y 11:00 UTC, tiene que asegurarse de que los argumentos `start` y `end` son válidos según UTC, lo que significa lo siguiente:<br/><br/><ul><li>`start.getUTCHours` devuelve 9:00 UTC</li><li>`end.getUTCHours` devuelve 11:00 UTC</li></ul>|Si las horas de inicio y de finalización del formulario de cita son 9:00 UTC y 11:00 UTC, tiene que asegurarse de que los argumentos `start` y `end` son válidos según UTC, lo que significa lo siguiente:<br/><br/><ul><li>`start.getUTCHours` devuelve 9:00 UTC</li><li>`end.getUTCHours` devuelve 11:00 UTC</li></ul>

## Métodos auxiliares para escenarios relacionados con fechas


Como se describe en las secciones anteriores, como la "hora local" de un usuario en Outlook Web App o en OWA para dispositivos puede ser distinta de la hora local en un cliente avanzado de Outlook, pero el objeto **Date** de JavaScript solo permite convertir a la zona horaria del equipo cliente o a UTC, la API de JavaScript para Office proporciona dos métodos auxiliares: [Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) y [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md). <br/><br/>
Estos métodos auxiliares solucionan cualquier necesidad para controlar la fecha o la hora de forma distinta para los dos escenarios siguientes relacionados con fechas, en un cliente avanzado de Outlook, Outlook Web App y OWA para dispositivos, lo que refuerza la "escritura de una sola vez" para los diferentes clientes del complemento.


### Escenario A: visualización de la hora de creación o modificación del elemento

Si quiere que en la interfaz de usuario se muestre la hora de creación del elemento (**Item.dateTimeCreated**) o la hora de modificación (**Item.dateTimeModified**), use en primer lugar **convertToLocalClientTime** para convertir el objeto **Date** proporcionado por estas propiedades para obtener una representación del diccionario en la hora local correcta. Después, muestre los componentes de la fecha del diccionario. Este es un ejemplo del escenario:


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook web app or OWA for Devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Tenga presente que  **convertToLocalClientTime** tiene en cuenta la diferencia entre un cliente enriquecido de Outlook, y Outlook Web App o bien OWA para dispositivos:


- Si  **convertToLocalClientTime** detecta que el host actual es un cliente enriquecido, el método convierte la representación **Date** a una representación de diccionario en la misma zona horaria que la del equipo cliente y coherente con el resto de la interfaz de usuario del cliente enriquecido.
    
- Si  **convertToLocalClientTime** detecta que el host actual es Outlook Web App u OWA para dispositivos, el método convierte la representación **Date** correcta según UTC a un formato de diccionario propio de la zona horaria de EAC y coherente con el resto de la interfaz de usuario de Outlook Web App o de OWA para dispositivos.
    

### Escenario B: visualización de las fechas de inicio y de finalización en un formulario de cita nuevo

Si obtiene como entrada diferentes partes de un valor de fecha y hora representado en la hora local y quiere proporcionar este valor de entrada de diccionario como una hora de inicio o finalización en un formulario de cita, use primero el método auxiliar **convertToUtcClientTime** para convertir el valor del diccionario en un objeto **Date** válido según UTC.<br/><br/>En el ejemplo siguiente, se supone que `myLocalDictionaryStartDate` y `myLocalDictionaryEndDate` son valores de fecha y hora en formato de diccionario que obtuvo del usuario. Estos valores se basan en la hora local y dependen de la aplicación host.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Los valores resultantes, `myUTCCorrectStartDate` y `myUTCCorrectEndDate`, son válidos según UTC. Después, pase estos objetos **Date** como argumentos para los parámetros _Start_ y _End_ del método **Mailbox.displayNewAppointmentForm** para mostrar el nuevo formulario de cita.<br/><br/>
Tenga presente que **convertToUtcClientTime** tiene en cuenta la diferencia entre un cliente avanzado de Outlook y Outlook Web App o OWA para dispositivos:


- Si  **convertToUtcClientTime** detecta que el host actual es un cliente enriquecido de Outlook, el método convierte simplemente la representación de diccionario en un objeto **Date**. Este objeto  **Date** es correcto en términos de hora UTC, tal como espera el método **displayNewAppointmentForm**.
    
- Si  **convertToUtcClientTime** detecta que el host actual es Outlook Web App u OWA para dispositivos, el método convierte el formato de diccionario de los valores de fecha y hora expresados en la zona horaria de EAC al objeto **Date**. Este objeto  **Date** es correcto en términos de hora UTC, tal como espera el método **displayNewAppointmentForm**.
    

## Recursos adicionales



- [Implementar e instalar complementos de Outlook para probarlos](../outlook/testing-and-tips.md)
    


