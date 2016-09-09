
# API de complementos de Outlook

Para usar las API en el complemento de Outlook, debe especificar la ubicación de la biblioteca Office.js, el conjunto de requisitos, el esquema y los permisos.

## Biblioteca Office.js

Para interactuar con la API de complemento de Outlook es necesario usar las API de JavaScript en Office.js. La CDN para la biblioteca es _https://appsforoffice.microsoft.com/lib/1/hosted/Office.js_. Los complementos enviados a la Tienda Office tienen que hacer referencia a Office.js con esta CDN, no pueden usar una referencia local. 

Declare la CDN en la etiqueta **head** de la página web (archivo .html, .aspx o .php) que implemente la interfaz de usuario del complemento, en el atributo **src** de la etiqueta **script**:


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

A medida que agregamos nuevas API, la dirección URL de Office.js seguirá siendo la misma. Cambiaremos la versión en la dirección URL solo si se interrumpe un comportamiento existente de la API.

> **Importante:** Al desarrollar un complemento para una aplicación host de Office, haga referencia a la API de JavaScript para Office desde dentro de la sección `<head>` de la página. Esto garantiza que la API se inicializa por completo antes de los elementos body. Los hosts de Office necesitan que los complementos se inicialicen 5 segundos después de la activación. Al superar este umbral, el complemento no responde y se muestra un mensaje de error al usuario.  

## Conjuntos de requisitos

Todas las API de Outlook pertenecen al conjunto de requisitos del buzón. El conjunto de requisitos del buzón tiene versiones y cada nuevo conjunto de API que publicamos pertenece a una versión posterior del conjunto. No todos los clientes de Outlook admitirán el conjunto más reciente de API cuando lo publiquemos. Pero, si un cliente de Outlook declara la compatibilidad con un conjunto de requisitos, será compatible con todas las API de ese conjunto de requisitos. 

Para controlar los clientes de Outlook donde aparece el complemento, especifique una versión mínima de conjunto de requisitos en el manifiesto. Por ejemplo, si especifica la versión 1.3 del conjunto de requisitos, el complemento no se mostrará en los clientes de Outlook que no sean compatibles con una versión mínima de 1.3. 

Especificar un requisito no limita el complemento a las API de esa versión. Si en el complemento se especifica como conjunto de requisitos la versión 1.1, pero se ejecuta en un cliente que es compatible con la versión 1.3, el complemento puede usar las API de la versión 1.3. El conjunto de requisitos solo controla los clientes de Outlook donde se mostrará el complemento.

Para comprobar la disponibilidad de las API de un conjunto de requisitos superior al especificado en el manifiesto, puede usar la técnica de JavaScript estándar:


```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> **Nota:** Estos controles son necesarios para cualquier API que esté en la versión del conjunto de requisitos especificada en el manifiesto.

Especifique el conjunto de requisitos mínimo que admita el conjunto crítico de las API de su escenario, sin el cual las características críticas del complemento no funcionarán. Especifique el conjunto de requisitos en el manifiesto de los elementos **Requirements**, **Sets** y **Set**. Para más información, vea [Manifiestos de complementos de Outlook](../outlook/manifests/manifests.md) e [Información sobre los conjuntos de requisitos de la API de Outlook](..\..\reference\outlook\tutorial-api-requirement-sets.md).

El elemento **Methods** no se aplica a los complementos de Outlook, por lo que no se puede declarar la compatibilidad para métodos específicos.


## Permisos

El complemento necesita los permisos adecuados para usar las API que necesita. Hay cuatro niveles de permisos, que se resumen a continuación. Para más información, vea [Información sobre los permisos del complemento de Outlook](../outlook/understanding-outlook-add-in-permissions.md).


|**Nivel de permisos**|**Descripción**|
|:-----|:-----|
|Restringido|Permite el uso de entidades, pero no de expresiones regulares.|
|Elemento de lectura|Además de lo que se permite en _Restringido_, permite lo siguiente:<ul><li>expresiones regulares</li><li>acceso de lectura a la API del complemento de Outlook</li><li>obtener las propiedades del elemento y el token de devolución de llamada</li></ul>|
|Lectura/escritura|Además de lo que se permite en _Leer elemento_, permite lo siguiente:<ul><li>acceso completo a la API del complemento de Outlook, excepto <b>makeEwsRequestAsync</b></li><li>configurar las propiedades del elemento</li></ul>|
|Buzón de lectura y escritura|Además de lo que se permite en _Lectura/escritura_, permite lo siguiente:<ul><li>crear, leer y escribir en elementos y carpetas</li><li>enviar elementos</li><li>realizar llamadas a [makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md#makeewsrequestasyncdata-callback-usercontext)</li></ul>|
En general, tiene que especificar el permiso mínimo necesario para el complemento. Los permisos se declaran en el elemento **Permissions** del manifiesto. Para más información, vea [Manifiestos de complementos de Outlook](../outlook/manifests/manifests.md). Para obtener información sobre problemas de seguridad, vea [Privacidad, permisos y seguridad para los complementos de Outlook](../outlook/../../docs/develop/privacy-and-security.md).


## Recursos adicionales

- [Manifiestos de complementos de Outlook](../outlook/manifests/manifests.md)

- [Entender los conjuntos de requisitos de la API de Outlook](../../reference/outlook/tutorial-api-requirement-sets.md)
    
- [Privacidad, permisos y seguridad para los complementos de Outlook](../outlook/../../docs/develop/privacy-and-security.md)
    
