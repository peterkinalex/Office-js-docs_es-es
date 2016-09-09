
# Elemento Rule
Especifica las reglas de activación que deberían evaluarse para este complemento de correo.

 **Tipo de complemento:** correo


## Sintaxis:

 **Regla ItemIs**: define una regla que evalúa en verdadero si el elemento seleccionado es del tipo especificado.


```XML
<Rule xsi:type="ItemIs" 
   ItemType= ["Appointment" | "Message"]
   FormType=["Read" | "Edit" | "ReadOrEdit"] 
   ItemClass = "string " 
   IncludeSubClasses=["true" | "false"] />
```

 **Regla ItemHasAttachment**: define una regla que evalúa en verdadero si el elemento contiene datos adjuntos.




```XML
<Rule xsi:type="ItemHasAttachment"  />
```

 **ItemHasKnownEntity**: define una regla que evalúa en verdadero si el elemento contiene texto del tipo de entidad especificado en el asunto o en el cuerpo.




```XML
<Rule xsi:type="ItemHasKnownEntity" 
  EntityType=["MeetingSuggestion" | "TaskSuggestion" |"Address" | "Url" | "PhoneNumber" | "EmailAddress" | "Contact" ]
  RegExFilter="string "
  FilterName="string "
  IgnoreCase=["true | false"]/>
```

 **Regla ItemHasRegularExpressionMatch**: define una regla que evalúa en verdadero si se encuentra una coincidencia para la expresión regular especificada en la propiedad indicada del elemento.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="string " 
    RegExValue="string " 
    PropertyName=["Subject" | "BodyAsPlaintext" | "BodyAsHtml" | "SenderSTMPAddress"]
    IgnoreCase=["true" | "false"]
/>
```

 **Regla RuleCollection**: define una colección de reglas y el operador lógico que se debe utilizar cuando se evalúen.




```XML
<Rule xsi:type="RuleCollection" Mode=["And" | "Or"]>
   ...
</Rule>
```


## Forma parte de:

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## Atributos:

 **Atributos de la regla ItemIs**



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|ItemType|ItemType (cadena)|necesario|Especifica el tipo de elemento con el que se debe coincidir. Puede ser uno de las siguientes:

|**ItemType**|**Corresponding ItemClass**|
|:-----|:-----|
|Cita|IPM.Appointment|
|Mensaje(1)|Incluye mensajes de correo electrónico, convocatorias de reunión, respuestas y cancelaciones.|
|
|FormType|ItemFormType (cadena)|necesario|Especifica si la aplicación debe aparecer en el formulario de edición o lectura para el elemento. Puede ser uno de las siguientes:|

|**FormType**|**Descripción**|
|:-----|:-----|
|Lectura|Especifica activar el complemento de correo solo en formularios de lectura (del valor **ItemType** especificado).|
|Editar|Especifica activar el complemento de correo solo en formularios de redacción (del valor **ItemType** especificado).|
|ReadOrEdit|Especifica activar el complemento de correo en formularios de lectura y de redacción (del valor **ItemType** especificado).|
|ItemClass|string|opcional|Especifica la clase de mensaje personalizada con la que se debe coincidir. Para obtener más información, consulte [Activar un complemento de correo de Outlook para una clase de mensaje específica](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx).|
|IncludeSubClasses|booleano|opcional|Especifica si la regla debería evaluar en verdadero si el elemento es de una subclase de la clase del mensaje especificada; el valor predeterminado es falso.|


(1) Los siguientes valores son las clases de mensajes correspondientes: IPM.NoteIPM.Schedule.Meeting.RequestIPM.Schedule.Meeting.NegIPM.Schedule.Meeting.PosIPM.Schedule.Meeting.TentIPM.Schedule.Meeting.Canceled.

 **Atributos de la regla ItemHasAttachment**

Ninguno.

 **Atributos de la regla ItemHasKnownEntity**



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|EntityType|KnownEntityType (cadena)|necesario|Especifica el tipo de entidad que se tiene que encontrar para que la regla evalúe en verdadero. Puede ser uno de los siguientes:

|**KnownEntityType**|**Descripción**|
|:-----|:-----|
|MeetingSuggestion|Texto que se identifica con el reconocimiento de plantillas para hacer referencia a un evento o a una reunión.|
|TaskSuggestion| Texto que se identifica con el reconocimiento de plantillas para incluir una frase que se puede accionar.|
|Address|Texto que se identifica con el reconocimiento de plantillas para hacer referencia a una dirección postal de los Estados Unidos.|
|Url|Texto que se identifica con el reconocimiento de plantillas para incluir un nombre de archivo o una dirección web URL.|
|PhoneNumber| Una serie de dígitos que se identifica con el reconocimiento de plantillas como un número de teléfono en Norteamérica.|
|EmailAddress|Texto que se identifica con el reconocimiento de plantillas para incluir una dirección de correo electrónico con formato SMTP.|
|Contact|Texto que se identifica con el reconocimiento de plantillas para incluir información de contacto.|
|RegExFilter|string|opcional|Especifica una expresión regular que se debe ejecutar con esta entidad para su activación.|
|FilterName|string|opcional|Especifica el nombre del filtro de expresión regular, de modo que después sea posible hacerle referencia en el código de su complemento.|
|IgnoreCase|booleano|opcional|Especifica que se ignoren las mayúsculas y minúsculas cuando se ejecute la expresión regular especificada por el atributo **RegExFilter**.|
 **Atributos de la regla ItemHasRegularExpressionMatch**



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|RegExName|string|necesario|Especifica el nombre de una expresión regular para que pueda hacer referencia a dicha expresión en el código de su complemento.|
|RegExValue|string|necesario|Especifica la expresión regular que se evaluará para determinar si se debe mostrar el complemento de correo. |
|PropertyName|PropertyName (cadena)|necesario|Especifica el nombre de la propiedad que contra la que se evaluará la expresión regular. Puede ser uno de las siguientes:

|**PropertyName**|**Descripción**|
|:-----|:-----|
|Tema|Evalúa la expresión regular según el asunto del elemento.|
|BodyAsPlaintext|Evalúa la expresión regular según el cuerpo del elemento en texto sin formato.|
|BodyAsHtml|Evalúa la expresión regular según el cuerpo del elemento si el cuerpo está disponible en HTML.|
|SenderSTMPAddress|Evalúa la expresión regular según la dirección SMTP del remitente del elemento.|
|IgnoreCase|booleano|opcional|Especifica que se ignoren las mayúsculas y minúsculas cuando se ejecute la expresión regular.|
 **Atributos de la regla de RuleCollection**



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Moda|string|necesario|Especifica el operador lógico que se usará al evaluar esta colección de reglas. Puede ser: "And" o "Or".|

## Recursos adicionales



- 
  [Activar un complemento de correo de Outlook para una clase de mensaje específica](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx) y [Reglas de activación para los complementos de Outlook](../../docs/outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins)
    
- [Coincidencia de cadenas en un elemento de Outlook como entidades conocidas](../../docs/outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [Usar las reglas de activación de las expresiones regulares para mostrar un complemento de Outlook](../../docs/outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
