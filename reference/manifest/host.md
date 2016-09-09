
# Elemento Host
Especifica el tipo de aplicación host de Office que admite su complemento de Office.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## Sintaxis:


```XML
<Host Name= ["Document" | "Database" | "Mailbox" | "Presentation" | "Project" | "Workbook"] />
```


## Atributos



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Nombre|string|necesario|El nombre del tipo de aplicación host de Office.|

## Observaciones

Puede especificar los siguientes valores en el atributo **Name** de un elemento **Host**. Cada valor se asigna al conjunto de una o varias aplicaciones host de Office admitidas por su complemento.



|**Nombre**|**Aplicaciones host de Office**|
|:-----|:-----|
| `"Document"`|Word, Word Online, Word en iPad|
| `"Database"`|aplicaciones web de Access|
| `"Mailbox"`|Outlook, Outlook Web App, OWA para dispositivos|
| `"Notebook"`|OneNote Online|
| `"Presentation"`|PowerPoint, PowerPoint Online, PowerPoint en iPad|
| `"Project"`|Project|
| `"Workbook"`|Excel, Excel Online, Excel en iPad|

## Observaciones

Para obtener más información sobre cómo especificar la compatibilidad del host, consulte [Especificar los requisitos de la API y del host de Office](../../docs/overview/specify-office-hosts-and-api-requirements.md).

