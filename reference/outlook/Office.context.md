

# <a name="context"></a>contexto

## <a name="officeofficemdcontext"></a>[Office](Office.md).context

El espacio de nombres de Office .context proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres Office.context, vea [Referencia de Office.context de referencia de la API compartida](../shared/office.context.md).

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="namespaces"></a>Espacios de nombres

[mailbox](Office.context.mailbox.md): Proporciona acceso al modelo de objetos del complemento de Outlook para Microsoft Outlook y Microsoft Outlook en la web.

### <a name="members"></a>Miembros

####  <a name="displaylanguage-string"></a>displayLanguage :String

Obtiene la configuración local (de idioma) en un formato de etiqueta de idioma RFC 1766 especificado por el usuario para la interfaz de usuario de la aplicación host de Office.

El valor `displayLanguage` refleja la configuración de **Mostrar idioma** actual que se ha especificado desde **Archivo > Opciones > Idioma**, en la aplicación host de Office.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

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

####  <a name="officetheme-object"></a>officeTheme :Object

Proporciona acceso a las propiedades de los colores para temas de Office.

> **Nota:** Este miembro no se admite en Outlook para iOS ni en Outlook para Android.

El uso de los colores del tema de Office le permite coordinar la combinación de colores del complemento con el tema actual de Office seleccionado por el usuario mediante **Archivo > Cuenta de Office > Interfaz de usuario Tema de Office**, que se aplica a todas las aplicaciones host de Office. El uso de colores del tema de Office es idóneo para los complementos de correo y panel de tareas.

##### <a name="type"></a>Tipo:

*   Objeto

##### <a name="properties"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`bodyBackgroundColor`| String|Obtiene el color de fondo del cuerpo del tema de Office como un tríptico de color hexadecimal.|
|`bodyForegroundColor`| String|Obtiene el color de primer plano del cuerpo del tema de Office como un tríptico de color hexadecimal.|
|`controlBackgroundColor`| String|Obtiene el color de fondo del control del tema de Office como un tríptico de color hexadecimal.|
|`controlForegroundColor`| String|Obtiene el color del control del cuerpo del tema de Office como un tríptico de color hexadecimal.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

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

####  <a name="roamingsettings-roamingsettingsroamingsettingsmd"></a>roamingSettings :[RoamingSettings](RoamingSettings.md)

Obtiene un objeto que representa la configuración o el estado personalizado de un complemento de correo que se guardó en el buzón de un usuario.

El objeto `RoamingSettings` le permite almacenar y tener acceso a datos para un complemento de correo almacenado en el buzón de un usuario, de forma que esté disponible para ese complemento cuando se ejecute desde cualquier aplicación de cliente host usada para tener acceso a ese buzón.

##### <a name="type"></a>Tipo:

*   [RoamingSettings](RoamingSettings.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|
