
# Propiedad Context.mailbox
Obtiene el objeto **mailbox** que proporciona acceso a los miembros de la API específicos para los complementos de Office.

|||
|:-----|:-----|
|**Hosts:**|Outlook|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Buzón|
|**Modificado por última vez en**|1,0|

```js
var outlookOm = Office.context.mailbox;
```


## Valor devuelto

El objeto [mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx).


## Ejemplo

La línea de código siguiente obtiene acceso al objeto [item](http://msdn.microsoft.com/library/ad288df1-3ca2-474c-bea4-c51f46e6fc43%28Office.15%29.aspx) de la API de JavaScript para Office.


```js
// Access the Item object.
var item = Office.context.mailbox.item;

```




## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Outlook para Mac**|
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
