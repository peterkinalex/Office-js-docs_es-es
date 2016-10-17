
# <a name="sourcelocation-element"></a>Elemento SourceLocation
Especifica las ubicaciones del código fuente para su complemento de Office como una dirección URL de entre 1 y 2018 caracteres. La ubicación de origen debe ser una dirección HTTPS, no una ruta de acceso de archivo.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax:"></a>Sintaxis:


```XML
<SourceLocation DefaultValue="string " />
```


## <a name="contained-in:"></a>Forma parte de:

[DefaultSettings](../../reference/manifest/defaultsettings.md) (complementos de contenido y de panel de tareas)

[FormSettings](../../reference/manifest/formsettings.md) (complementos de correo)


## <a name="can-contain:"></a>Puede contener:

[Override](../../reference/manifest/override.md)


## <a name="attributes"></a>Atributos



|**Atributo**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultValue|Dirección URL|necesario|Especifica el valor predeterminado de esta opción para la configuración regional especificada en el elemento [DefaultLocale](../../reference/manifest/defaultlocale.md).|
