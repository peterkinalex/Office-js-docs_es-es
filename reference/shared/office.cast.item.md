
# Propiedad Office.cast.item
Proporciona IntelliSense específico para mensajes y citas en modo de redacción o lectura.

|||
|:-----|:-----|
|**Hosts:**|Outlook|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Buzón|
|**Modificado por última vez en**|1,0|



|||
|:-----|:-----|
|**Modos de Outlook aplicables**|Tiempo de diseño solo en Visual Studio|

```js
Office.cast.item.toAppointmentCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointmentRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointment(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessage(Office.context.mailbox.item);
```


## Valor devuelto

Conjunto de métodos que le permiten seleccionar el IntelliSense apropiado para su complemento de Outlook.


## Comentarios

Esta propiedad y sus métodos solo son compatibles con el uso de IntelliSense para desarrollar complementos de Outlook en Visual Studio. El resto de las herramientas de desarrollo no resultan afectadas.

Los métodos **Office.cast.item** se usan en tiempo de diseño en Visual Studio para proporcionar IntelliSense específico para la propiedad **Office.context.mailbox.item**. Si usa el método **toAppointmentCompose**, por ejemplo, IntelliSense solo le mostrará las propiedades y los métodos **Appointment** aplicables en modo de redacción.

En tiempo de ejecución, los métodos **Office.cast.item** no tendrán efecto alguno en su complemento de Outlook.


## Ejemplo

En el ejemplo siguiente se usa el método **toMessageCompose** para convertir la propiedad **Office.context.mailbox.item** de forma que solo mostrará IntelliSense para el objeto **Message** en modo de redacción. Después de la conversión, la variable `message` solo mostrará IntelliSense para los métodos y propiedades que se pueden usar en el modo de redacción.


```js
var message = Office.cast.item.toMessageCompose(Office.context.mailbox.item);

```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||Office para escritorio de Windows|Office Online (en el explorador)|Outlook para Mac|
|:-----|:-----|:-----|:-----|
|**Outlook**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Buzón|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1,0|Agregado|
