
# Elemento SourceLocation
Especifica las ubicaciones del código fuente para su complemento de Office como una dirección URL de entre 1 y 2018 caracteres. La ubicación de origen debe ser una dirección HTTPS, no una ruta de acceso de archivo.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## Sintaxis:


```XML
<SourceLocation DefaultValue="string " />
```


## Forma parte de:

[DefaultSettings](../../reference/manifest/defaultsettings.md) (complementos de contenido y de panel de tareas)

[FormSettings](../../reference/manifest/formsettings.md) (complementos de correo)


## Puede contener:

[Override](../../reference/manifest/override.md)


## Atributos



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultValue|Dirección URL|necesario|Especifica el valor predeterminado de esta opción para la configuración regional especificada en el elemento [DefaultLocale](../../reference/manifest/defaultlocale.md).|
