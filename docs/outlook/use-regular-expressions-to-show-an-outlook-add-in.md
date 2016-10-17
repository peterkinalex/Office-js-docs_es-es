
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>Usar las reglas de activación de las expresiones regulares para mostrar un complemento de Outlook

Puede especificar reglas de expresiones regulares para que un complemento de Outlook se active en escenarios de lectura. Cuando el usuario vea un mensaje o una cita en el inspector o en el panel de lectura, Outlook evaluará las reglas de expresión regular para determinar si tiene que activar el complemento de correo. Outlook no evalúa estas reglas mientras el usuario redacta un elemento. Existen otros escenarios en los que Outlook no activa complementos (por ejemplo, si los elementos se encuentran protegidos con Information Rights Management [IRM] o en la carpeta Correo no deseado). Para más información, vea [Reglas de activación para complementos de Outlook](../outlook/manifests/activation-rules.md).

Si lo desea, puede especificar una expresión regular como parte de una regla [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) o [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) en el archivo XML de manifiesto del complemento. Outlook evalúa las expresiones regulares a partir de las reglas del intérprete de JavaScript que el explorador ha usado en el equipo cliente y admite la misma lista de caracteres especiales (en la tabla siguiente) que los procesadores XML. Para usar estos caracteres en una expresión regular, especifique la secuencia de escape para el carácter correspondiente conforme a la tabla siguiente.



|**Carácter**|**Descripción**|**Secuencia de escape que debe usarse**|
|:-----|:-----|:-----|
|"|Comilla doble|&amp;quot;|
|&amp;|Y comercial|&amp;amp;|
|'|Apóstrofo|&amp;apos;|
|<|Signo de menor que|&amp;lt;|
|>|Signo de mayor que|&amp;gt;|

## <a name="itemhasregularexpressionmatch-rule"></a>Regla ItemHasRegularExpressionMatch


Una regla  **ItemHasRegularExpressionMatch** es útil para controlar la activación de un complemento en función de valores específicos de una propiedad admitida. La regla **ItemHasRegularExpressionMatch** tiene los siguientes atributos.



|**Nombre del atributo**|**Descripción**|
|:-----|:-----|
|**RegExName**|Especifica el nombre de una expresión regular para que pueda hacer referencia a dicha expresión en el código del complemento.|
|**RegExValue**|Especifica la expresión regular que se evaluará para determinar si se debe mostrar el complemento.|
|**PropertyName**|Especifica el nombre la propiedad con respecto a la cual se evaluará la expresión regular. Los valores permitidos son  **BodyAsHTML**,  **BodyAsPlaintext**,  **SenderSMTPAddress** y **Subject**. Si especifica  **BodyAsHTML**, Outlook aplica la expresión regular solo si el cuerpo del elemento es HTML; de lo contrario, Outlook no devuelve ningún resultado para dicha expresión regular. Dado que las citas siempre se guardan en formato de texto enriquecido, una expresión regular que especifica  **BodyAsHTML** no coincide con ninguna cadena del cuerpo de elementos de cita.Si especifica  **BodyAsPlaintext**, Outlook siempre aplica la expresión regular al cuerpo del elemento.|
|**IgnoreCase**|Especifica si debe distinguirse entre mayúsculas y minúsculas al buscar resultados con la expresión regular especificada por **RegExName**.|

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>Procedimientos recomendados para usar expresiones regulares en reglas

Preste especial atención a lo siguiente cuando use expresiones regulares:


- Si especifica una regla  **ItemHasRegularExpressionMatch** en el cuerpo de un elemento, la expresión regular debe seguir filtrando el cuerpo y no tratar de devolver todo el cuerpo del elemento. El uso de una expresión regular como `.*` para tratar de obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.
    
- El cuerpo de texto sin formato devuelto en un explorador puede ser ligeramente diferente en otro. Si usa una regla [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) con **BodyAsPlaintext** como atributo **PropertyName**, pruebe la expresión regular en todos los exploradores compatibles con su complemento.
    
    Dado que los distintos exploradores usan diferentes formas de obtener el cuerpo del texto de un elemento seleccionado, debe asegurarse de que la expresión regular es compatible con las diferencias sutiles que pueden devolverse como parte del texto del cuerpo. Por ejemplo, algunos exploradores, como Internet Explorer 9 usan la propiedad **innerText** del DOM, y otros como Firefox usan el método **.textContent()** para obtener el cuerpo del texto de un elemento. Además, cada explorador puede devolver los saltos de línea de manera diferente: un salto de línea es "\r\n" en Internet Explorer y "\n" en Firefox y Chrome. Para obtener más información, consulte [Compatibilidad de DOM del W3C - HTML](http://www.quirksmode.org/dom/w3c_html.mdl#t07).
    
- El cuerpo HTML de un elemento difiere ligeramente entre un cliente enriquecido de Outlook y Outlook Web App o OWA para dispositivos. Defina sus expresiones regulares con cuidado. A modo de ejemplo, considere la siguiente expresión regular usada en una regla  **ItemHasRegularExpressionMatch** con **BodyAsHTML** como valor del atributo **PropertyName**:
    
```
      http.*\.contoso\.com
```


    A rule with this regular expression would match the string "http-equiv="Content-Type" which exists in the HTML body of an item in an Outlook rich client, as part of the following  **META** tag:
    

```HTML
      <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii">
```


La misma regla no devuelve este resultado en Outlook Web App y OWA para dispositivos porque en el cuerpo HTML de estos hosts no se incluye la etiqueta **META**. Esto puede afectar a la activación correcta del complemento en los distintos clientes de Outlook. En este ejemplo, use en su lugar la siguiente expresión regular:
    

```
      http://.*\.contoso\.com/
```

- En función de la aplicación host, el tipo de dispositivo o la propiedad a la que se aplica una expresión regular, existen otros procedimientos recomendados y límites para cada uno de los hosts que debe tener en cuenta al diseñar expresiones regulares como reglas de activación. Vea [Límites de activación y API de JavaScript para complementos de Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md) para más información.
    

### <a name="examples"></a>Ejemplos

La siguiente regla  **ItemHasRegularExpressionMatch** activa el complemento siempre que la dirección de correo electrónico SMTP del remitente coincida con "@contoso", independientemente de si los caracteres están en mayúsculas o minúsculas.


```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]" 
    PropertyName="SenderSMTPAddress"
/>
```

La siguiente es otra forma de especificar la misma expresión regular con el atributo  **IgnoreCase**.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@contoso" 
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

La siguiente regla  **ItemHasRegularExpressionMatch** activa el complemento siempre que se incluya el símbolo de un valor en el cuerpo del elemento actual.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    PropertyName="BodyAsPlaintext" 
    RegExName="TickerSymbols" 
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```


## <a name="itemhasknownentity-rule"></a>Regla ItemHasKnownEntity


Una regla  **ItemHasKnownEntity** activa un complemento en función de la existencia de una entidad en el asunto o cuerpo del elemento seleccionado. El tipo [KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) define las entidades admitidas. La aplicación de una expresión regular a una regla **ItemHasKnownEntity** es conveniente en aquellos casos en los que la activación se basa en un subconjunto de valores para una entidad (por ejemplo, un conjunto específico de direcciones URL o números de teléfono con un determinado prefijo regional).


 >
  **Nota**  Outlook puede extraer cadenas de entidades solo en inglés, independientemente de la configuración regional predeterminada que se especifique en el manifiesto. Solo los mensajes, y no las citas, admiten el tipo de entidad  **MeetingSuggestion**.No pueden extraer entidades de elementos de la carpeta Elementos enviados ni tampoco usar una regla [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) para activar un complemento para elementos de dicha carpeta.

La regla  **ItemHasKnownEntity** admite los atributos de la siguiente tabla. Tenga en cuenta que, si bien la especificación de una expresión regular es opcional en una regla **ItemHasKnownEntity**, si opta por usar una expresión regular como filtro de entidades, deberá especificar tanto el atributo  **RegExFilter** como el atributo **FilterName**.



|**Nombre del atributo**|**Descripción**|
|:-----|:-----|
|**EntityType**|Especifica el tipo de entidad que debe encontrarse para que la regla se evalúe como  **true**. Use varias reglas para especificar varios tipos de entidad.|
|**RegExFilter**|Especifica una expresión regular que filtra aún más las instancias de la entidad especificada por **EntityType**.|
|**FilterName**|Especifica el nombre de la expresión regular especificada por **RegExFilter**, de tal modo que es posible referirse a ella posteriormente con código.|
|**IgnoreCase**|Especifica si debe distinguirse entre mayúsculas y minúsculas al buscar resultados con la expresión regular especificada por **RegExFilter**.|

### <a name="examples"></a>Ejemplos

La siguiente regla  **ItemHasKnownEntity** activa el complemento siempre que haya una dirección URL en el asunto o cuerpo del elemento actual y la dirección URL contenga la cadena "youtube", independientemente de si está en mayúsculas o minúsculas.


```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```


## <a name="using-regular-expression-results-in-code"></a>Usar los resultados de expresiones regulares en el código


Puede obtener resultados para una expresión regular usando los siguientes métodos en el elemento actual:


- [getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md) devuelve resultados en el elemento actual para todas las expresiones regulares que se especifican en las reglas **ItemHasRegularExpressionMatch** e **ItemHasKnownEntity** del complemento.
    
- [getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md) devuelve resultados en el elemento actual para la expresión regular identificada que se especifica en una regla **ItemHasRegularExpressionMatch** del complemento.
    
- [getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md) devuelve instancias enteras de entidades que contengan resultados para la expresión regular identificada que se especifica en una regla **ItemHasKnownEntity** del complemento.
    
Cuando se evalúan las expresiones regulares, se devuelven los resultados al complemento en un objeto de matriz. En el caso de  **getRegExMatches**, dicho objeto tiene el identificador del nombre de la expresión regular. 


 >**Nota**  Un cliente enriquecido de Outlook no devuelve resultados en ningún orden específico en la matriz. Además, no debe dar por sentado que el cliente enriquecido de Outlook va a devolver resultados en el mismo orden en esta matriz que en Outlook Web App o OWA para dispositivos, aunque ejecute el mismo complemento en cada uno de estos clientes en el mismo elemento del mismo buzón. Para conocer otras diferencias a la hora de procesar expresiones regulares entre un cliente enriquecido de Outlook y Outlook Web App o OWA para dispositivos, vea [Límites de activación y API de JavaScript para complementos de Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).


### <a name="examples"></a>Ejemplos

El ejemplo siguiente es una colección de reglas que contiene una regla  **ItemHasRegularExpressionMatch** con una expresión regular de nombre `videoURL`.


```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="VideoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="Body"/>
</Rule>
```

El ejemplo siguiente usa  **getRegExMatches** del elemento actual para establecer una variable `videos` para los resultados de la regla **ItemHasRegularExpressionMatch** precedente.




```
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

Varios resultados se almacenan como elementos de matriz en dicho objeto. El siguiente ejemplo de código muestra cómo iterar los resultados de una expresión regular de nombre  `reg1` para crear una cadena que se muestre como HTML.




```js
function initDialer() 
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = _Item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }
    myCell.innerHTML = myString;
}

```

El ejemplo siguiente muestra una regla  **ItemHasKnownEntity** que especifica la entidad **MeetingSuggestion** y una expresión regular de nombre `CampSuggestion`. Outlook activa el complemento si detecta que el elemento seleccionado contiene una sugerencia de reunión y el asunto o cuerpo contienen el término "WonderCamp".




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

El siguiente ejemplo de código usa  **getFilteredEntitiesByName(name)** del elemento actual para establecer una variable `suggestions` para obtener una matriz de las sugerencias de reunión detectadas para la regla **ItemHasKnownEntity** precedente.




```
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName(CampSuggestion);
```


## <a name="additional-resources"></a>Recursos adicionales



- [Crear complementos de Outlook para formularios de lectura](../outlook/read-scenario.md)
    
- [Reglas de activación para complementos de Outlook](../outlook/manifests/activation-rules.md)
    
- [Límites para la activación y API de JavaScript para complementos de Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [Coincidencia de cadenas en un elemento de Outlook como entidades conocidas](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- 
  [Procedimientos recomendados para expresiones regulares en .NET Framework](http://msdn.microsoft.com/en-us/library/618e5afb-3a97-440d-831a-70e4c526a51c%28Office.15%29.aspx)
    
