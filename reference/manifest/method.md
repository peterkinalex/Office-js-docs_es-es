
# Elemento Method
Especifica un método individual de la API de JavaScript para Office que su complemento de Office necesita para activarse.

 **Tipo de complemento:** Panel de tareas, contenido


## Sintaxis:


```XML
<Method Name="string "/>
```


## Forma parte de:

 _ [Métodos](../../reference/manifest/methods.md)_


## Atributos



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Nombre|string|necesario|Especifica el nombre del método necesario calificado con su objeto principal. Por ejemplo, para especificar el método **getSelectedDataAsync**, debe especificar `"Document.getSelectedDataAsync"`.|

## Observaciones

Los elementos **Methods** y **Method** no son compatibles con los complementos de correo. Para obtener más información acerca de los conjuntos de requisitos, consulte [Especificar los requisitos de la API y del host de Office](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_intro).


 >**Importante**  Debido a que no hay ningún método para especificar el requisito de versión mínima para los métodos individuales y para asegurarse de que haya un método disponible en tiempo de ejecución, debe usar también una declaración **if** al llamar al método del script de su complemento. Para obtener más información sobre cómo hacerlo, consulte [Información sobre la API de JavaScript para Office](../../docs/develop/understanding-the-javascript-api-for-office.md#HostAPISupport_UsingIfStatements).

