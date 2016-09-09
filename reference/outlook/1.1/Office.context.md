

# contexto

## [Office](Office.md). context

El espacio de nombres de Office .context proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres Office.context, vea [Referencia de Office.context de referencia de la API compartida](../../shared/office.context.md).


##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|Modo de Outlook aplicable| Redacción o lectura|

### Espacios de nombres

[mailbox](Office.context.mailbox.md): Proporciona acceso al modelo de objetos del complemento de Outlook para Microsoft Outlook y Microsoft Outlook en la web.

### Miembros

####  displayLanguage :String

Obtiene la configuración local (de idioma) en un formato de etiqueta de idioma RFC 1766 especificado por el usuario para la interfaz de usuario de la aplicación host de Office.

El valor `displayLanguage` refleja la configuración de **Mostrar idioma** actual que se ha especificado desde **Archivo > Opciones > Idioma**, en la aplicación host de Office.

##### Tipo:

*   String

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|Modo de Outlook aplicable| Redacción o lectura|

##### Ejemplo

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  roamingSettings :[RoamingSettings](RoamingSettings.md)

Obtiene un objeto que representa la configuración o el estado personalizado de un complemento de correo que se guardó en el buzón de un usuario.

El objeto `RoamingSettings` le permite almacenar y tener acceso a datos para un complemento de correo almacenado en el buzón de un usuario, de forma que esté disponible para ese complemento cuando se ejecute desde cualquier aplicación de cliente host usada para tener acceso a ese buzón.

##### Tipo:

*   [RoamingSettings](RoamingSettings.md)

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|
