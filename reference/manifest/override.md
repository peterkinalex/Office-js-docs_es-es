
# Elemento Override
Permite especificar el valor de una opción para una configuración regional adicional.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## Sintaxis:


```XML
<Override Locale="string " Value="string " />
```


## Forma parte de:


||
|:-----|
|[CitationText](../../reference/manifest/citationtext.md)|
|[Descripción](../../reference/manifest/description.md)|
|[DictionaryName](../../reference/manifest/dictionaryname.md)|
|[DictionaryHomePage](../../reference/manifest/dictionaryhomepage.md)|
|[DisplayName](../../reference/manifest/displayname.md)|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|
|[IconUrl](../../reference/manifest/iconurl.md)|
|[QueryUri](../../reference/manifest/queryuri.md)|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|
|[SupportUrl](../../reference/manifest/supporturl.md)|

## Atributos



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Configuración regional|string|necesario|Especifica el nombre de referencia cultural de la configuración regional para esta invalidación en el formato de etiqueta de lenguaje BCP 47, como en `"en-US"`.|
|Valor|string|necesario|Especifica el valor de la opción de configuración expresado para la configuración regional especificada.|

## Recursos adicionales



- [Localización de complementos para Office](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
