

# <a name="match-strings-in-an-outlook-item-as-well-known-entities"></a>Coincidencia de cadenas en un elemento de Outlook como entidades conocidas


Antes de enviar un mensaje o un elemento de convocatoria de reunión, Exchange Server analiza el contenido del elemento, identifica y marca determinadas cadenas en el asunto y en el cuerpo que son similares a entidades conocidas de Exchange (por ejemplo, direcciones de correo electrónico, números de teléfono y direcciones URL). Los mensajes y las convocatorias de reunión son entregadas por Exchange Server en una bandeja de entrada de Outlook con entidades conocidas marcadas. 

Si usa la API de JavaScript para Office, puede obtener estas cadenas que coinciden con entidades específicas conocidas para su procesamiento posterior. También puede especificar una entidad conocida en una regla en el manifiesto del complemento para que Outlook pueda activar el complemento cuando el usuario visualice un elemento que contenga coincidencias con la entidad. Después, puede extraer y realizar una acción en las coincidencias de la entidad. 

Es conveniente saber identificar o extraer estas instancias a partir de un mensaje o cita seleccionados. Por ejemplo, puede crear un servicio inverso para buscar teléfonos como un complemento de Outlook que extrae cadenas del asunto o el cuerpo del elemento que parezcan un número de teléfono, realiza una búsqueda inversa y muestra el propietario registrado de cada número de teléfono.

En este tema se presentan estas entidades conocidas, se muestran ejemplos de reglas de activación basadas en entidades conocidas y se explica cómo extraer coincidencias de entidad independientemente de que se hayan utilizado entidades en las reglas de activación.


## <a name="support-for-well-known-entities"></a>Compatibilidad con entidades conocidas


Exchange Server marca las entidades conocidas de un mensaje o un elemento de convocatoria de reunión después de que el remitente envíe el elemento y antes de que Exchange entregue el elemento al destinatario. Por lo tanto, solo se marcan los elementos que hayan pasado por el servicio de transporte en Exchange, y Outlook puede activar los complementos según estas marcas cuando el usuario los visualiza. Por el contrario, cuando el usuario está redactando un elemento o visualizando un elemento que se encuentra en la carpeta Elementos enviados, como el elemento no ha pasado a través del servicio de transporte, Outlook no puede activar los complementos según las entidades conocidas. 

De forma similar, tampoco se pueden extraer entidades conocidas de elementos que estén siendo redactados o que se encuentren en la carpeta Elementos enviados, ya que estos elementos no han pasado a través del servicio de transporte y no han sido marcados. Para más información sobre los tipos de elementos compatibles con la activación, vea [Reglas de activación para complementos de Outlook](../outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins).

En la tabla siguiente se muestran las entidades que admiten y reconocen Exchange Server y Outlook (de ahí que se denominen "entidades conocidas") y el tipo de objeto de una instancia de cada entidad. El reconocimiento del lenguaje natural de una cadena como una de estas entidades se basa en un modelo de aprendizaje formado sobre una gran cantidad de datos. Por ello, el reconocimiento no es determinista. Vea [Sugerencias para usar entidades conocidas](#tips-for-using-well-known-entities) si desea más información sobre las condiciones del reconocimiento.

 **Tabla 1: Entidades admitidas y sus tipos**



|**Tipo de entidad**|**Condiciones para el reconocimiento**|**Tipo de objeto**|
|:-----|:-----|:-----|
|**Address**|Direcciones postales de los Estados Unidos; por ejemplo: 1234 Main Street, Redmond, WA 07722.En general, para que se pueda reconocer una dirección, debe seguir la estructura de las direcciones postales estadounidenses e incluir la mayoría de los siguientes elementos: número de la calle, nombre de la calle, ciudad, estado y código postal. La dirección se puede especificar en una o varias líneas.|Objeto JavaScript **String**|
|**Contact**|Una referencia a la información de una persona como se reconoce en el lenguaje natural.El reconocimiento de un contacto depende del contexto. Por ejemplo, una firma al final del mensaje, o el nombre de una persona que aparezca cerca de alguno de estos elementos de información: un número de teléfono, una dirección, una dirección de correo electrónico o una URL.|Objeto [Contact](../../reference/outlook/simple-types.md)|
|**EmailAddress**|Direcciones de correo electrónico SMTP.|Objeto JavaScript **String**|
|**MeetingSuggestion**|Una referencia a un evento o encuentro. Por ejemplo, Exchange 2013 reconocería el siguiente texto como una sugerencia de encuentro:  _Quedamos mañana para comer_|Objeto [MeetingSuggestion](../../reference/outlook/simple-types.md)|
|**PhoneNumber**|Números de teléfono de Estados Unidos; por ejemplo:  _(235) 555-0110_|Objeto [PhoneNumber](../../reference/outlook/simple-types.md)|
|**TaskSuggestion**|Frases accionables de un mensaje de correo. Por ejemplo:  _Por favor, actualice la hoja de cálculo._|Objeto [TaskSuggestion](../../reference/outlook/simple-types.md)|
|**Url**|Una dirección web que especifique explícitamente la ubicación de la red y el identificador de un recurso web. Exchange Server no requiere el protocolo de acceso en la dirección web y no reconoce las direcciones URL incrustadas en el texto de vínculos como instancias de la entidad  **Url**. Exchange Server puede devolver resultados para los siguientes ejemplos: _www.youtube.com/user/officevideos_ _http://www.youtube.com/user/officevideos_|Objeto JavaScript  **String**|
La figura 1 describe cómo Exchange Server y Outlook admiten entidades conocidas para complementos y qué pueden hacer estos con las entidades conocidas. Consulte los temas sobre la [Recuperación de las entidades en su complemento](#retrieving-entities-in-your-add-in) y la [Activación de un complemento sobre la base de la existencia de una entidad](#activating-an-add-in-based-on-the-existence-of-an-entity) para obtener más detalles sobre cómo usar estas entidades.


**Figura 1: Cómo Exchange Server, Outlook y los complementos admiten entidades conocidas**

![Soporte y uso de entidades conocidas en aplicación de correo](../../images/mod_off15_mailapp_wellknownentities_curvedlines.png)


## <a name="permissions-to-extract-entities"></a>Permisos para extraer entidades


Para extraer entidades en su código JavaScript o para que su complemento se active según la existencia de determinadas entidades conocidas, asegúrese de solicitar los permisos adecuados en el manifiesto del complemento.

Al especificar el permiso restringido predeterminado, el complemento puede extraer la entidad  **Address**,  **MeetingSuggestion** o **TaskSuggestion**. Para extraer las entidades restantes, especifique el permiso de buzón de lectura y escritura, un elemento de lectura y escritura o un elemento de lectura. Para especificarlo en el manifiesto, use el elemento [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) y especifique el permiso correspondiente ( **Restricted**,  **ReadItem**,  **ReadWriteItem** o **ReadWriteMailbox**), como en el ejemplo siguiente:




```XML
<Permissions>ReadItem</Permissions>
```


## <a name="retrieving-entities-in-your-add-in"></a>Recuperación de las entidades en su complemento


Siempre que el cuerpo o el asunto del elemento que se está viendo contenga cadenas que Exchange y Outlook pueden identificar como entidades conocidas, estas instancias se encontrarán disponibles para los complementos (aunque los complementos no estén activados en función de estas entidades). Con el permiso adecuado, puede usar el método  **getEntities** o **getEntitiesByType** para recuperar entidades conocidas que se encuentran en el mensaje o la cita actual. El método **getEntities** devuelve una matriz de objetos [Entities](../../reference/outlook/simple-types.md) que contiene todas las entidades conocidas del elemento. Si está interesado en un tipo de entidades en particular, el método **getEntitiesByType** permite obtener una matriz que contiene únicamente las entidades que se especifican. La enumeración [EntityType](../../reference/outlook/Office.MailboxEnums.md) representa todos los tipos de entidades conocidas que se pueden extraer.

Después de llamar a  **getEntities**, puede usar la propiedad correspondiente del objeto  **Entities** para obtener una matriz de instancias de un tipo de entidad. Dependiendo del tipo de entidad, las instancias de la matriz pueden ser simplemente cadenas o pueden estar asignadas a determinados objetos. Como el ejemplo en la figura 1, para obtener direcciones en el elemento, obtenga acceso a la matriz que devuelve `getEntities().addresses[]`. La propiedad  **Entities.addresses** devuelve una matriz de cadenas que Outlook reconoce como direcciones postales. Del mismo modo, la propiedad **Entities.contacts** devuelve una matriz de objetos **Contact** que Outlook reconoce como información de contacto. La tabla 1 describe el tipo de objeto de una instancia de cada entidad admitida.

El ejemplo siguiente muestra cómo recuperar direcciones que se encuentren en un mensaje.




```
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities &amp;&amp; null != entities.addresses &amp;&amp; undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## <a name="activating-an-add-in-based-on-the-existence-of-an-entity"></a>Activación de un complemento sobre la base de la existencia de una entidad


Cuando se usan entidades conocidas, Outlook también puede activar el complemento en función de la existencia de uno o varios tipos de entidades en el asunto o el cuerpo del elemento que se está viendo. Para ello, debe especificar una regla  **ItemHasKnownEntity** en el manifiesto del complemento. El tipo sencillo [KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) representa los diferentes tipos de entidades conocidas compatibles con las reglas **ItemHasKnownEntity**. Cuando se active el complemento, también podrá recuperar las instancias de estas entidades si así lo desea, tal como se describe en la sección anterior [Recuperación de las entidades en su complemento](#retrieving-entities-in-your-add-in). 

Si lo prefiere, puede aplicar una expresión regular a una regla  **ItemHasKnownEntity**, como también a otras instancias de filtrado de una entidad, y hacer que Outlook active un complemento solo en un subconjunto de las instancias de la entidad. Por ejemplo, puede especificar un filtro para la entidad de nombre de la calle en un mensaje que contenga un código postal del estado de Washington que empiece por "98". Para aplicar un filtro en las instancias de la entidad, use los atributos  **RegExFilter** y **FilterName** en el elemento [Rule](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx) del tipo [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx).

Al igual que otras reglas de activación, puede especificar múltiples reglas para formar una colección de reglas para su complemento. En el siguiente ejemplo se aplica una operación "AND" en 2 reglas: una regla  **ItemIs** y una regla **ItemHasKnownEntity**. Esta colección de reglas activa el complemento siempre que el elemento actual sea un mensaje y Outlook reconozca una dirección en el asunto o el cuerpo de ese elemento.




```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

El ejemplo siguiente usa  **getEntitiesByType** del elemento actual para establecer una variable `addresses` para los resultados de la colección de reglas anterior.




```
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

El ejemplo de regla  **ItemHasKnownEntity** siguiente activa el complemento siempre que haya una dirección URL en el asunto o cuerpo del elemento actual y la dirección URL contenga la cadena "youtube", independientemente de si está en mayúsculas o minúsculas.




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

El ejemplo siguiente usa  **getFilteredEntitiesByName(name)** del elemento actual para establecer una variable `videos` para obtener una matriz de resultados que coincidan con la expresión regular en la regla **ItemHasKnownEntity** precedente.




```
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## <a name="tips-for-using-well-known-entities"></a>Sugerencias para usar entidades conocidas


Existen algunos hechos y límites relacionados con el uso de entidades conocidas en un complemento que debe conocer. Las consideraciones siguientes se aplican siempre que un complemento se encuentra activado mientras el usuario lee un elemento que contiene coincidencias de entidades conocidas, independientemente de si usa o no una regla  **ItemHasKnownEntity**:


1. Puede extraer cadenas que sean entidades conocidas solo si las cadenas están en inglés.
    
2. Puede extraer entidades conocidas de los primeros 2000 caracteres del cuerpo del elemento, pero no más allá de este límite. Este límite en el tamaño ayuda a equilibrar las funciones y el rendimiento, de manera que Exchange Server y Outlook no se estanquen durante el análisis y la identificación de instancias de entidades conocidas en mensajes y citas grandes. Tenga en cuenta que este límite es independiente del hecho de que el complemento especifique una regla  **ItemHasKnownEntity**. Si el complemento usa esta regla, tenga en cuenta también el límite de procesamiento de reglas en el elemento 2 siguiente para los clientes Outlook enriquecidos.
    
3. Se pueden extraer entidades de citas que sean reuniones organizadas por alguien que no sea el propietario del buzón de correo. No se pueden extraer entidades de elementos de calendario que no sean reuniones o que sean reuniones organizadas por el propietario del buzón.
    
4. Se pueden extraer entidades del tipo  **MeetingSuggestion** pero solamente de los mensajes, no de las citas.
    
5. También se pueden extraer las URL que existan explícitamente en el cuerpo del elemento, pero no las URL incluidas en el texto de los hipervínculos en el cuerpo del elemento HTML. Considere usar una regla  **ItemHasRegularExpressionMatch** en vez de las URL explícitas y las incluidas en hipervínculos. Especifique **BodyAsHTML** como _PropertyName_, y una expresión regular que coincida con las URL como  _RegExValue_.
    
6. No se pueden extraer entidades de elementos de la carpeta Elementos enviados.
    
Además, lo siguiente se aplica si se usa una regla [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) y podría afectar a los escenarios donde el complemento debería activarse:


1. Cuando se usa la regla  **ItemHasKnownEntity**, Outlook debería dar como resultado cadenas de entidades solo en inglés, independientemente de la configuración regional predeterminada que se especifique en el manifiesto.
    
2. Cuando el complemento se ejecute en un cliente Outlook enriquecido, Outlook debe aplicar la regla  **ItemHasKnownEntity** al primer megabyte del cuerpo de elemento y no al resto del cuerpo más allá de ese límite.
    
3. No se puede usar una regla  **ItemHasKnownEntity** para activar un complemento para elementos en la carpeta Elementos enviados.
    

## <a name="additional-resources"></a>Recursos adicionales



- [Crear complementos de Outlook para formularios de lectura](../outlook/read-scenario.md)
    
- [Extraer cadenas de entidad de un elemento de Outlook](../outlook/extract-entity-strings-from-an-item.md)
    
- [Reglas de activación para complementos de Outlook](../outlook/manifests/activation-rules.md)
    
- [Usar las reglas de activación de las expresiones regulares para mostrar un complemento de Outlook](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Comprender los permisos de los complementos de Outlook](../outlook/understanding-outlook-add-in-permissions.md)
    
