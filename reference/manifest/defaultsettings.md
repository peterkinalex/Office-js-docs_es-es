
# <a name="defaultsettings-element"></a>Elemento DefaultSettings
Especifica la ubicación del código fuente predeterminada y otras configuraciones predeterminadas de su complemento de contenido o de panel de tareas.

 **Tipo de complemento:** Panel de tareas, contenido


## <a name="syntax:"></a>Sintaxis:


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## <a name="contained-in:"></a>Forma parte de:

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain:"></a>Puede contener:



|**Elemento**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|x||x|
|[RequestedWidth](../../reference/manifest/requestedwidth.md)|x|||
|[RequestedHeight](../../reference/manifest/requestedheight.md)|x|||

## <a name="remarks"></a>Comentarios

La ubicación del origen y otros parámetros del elemento **DefaultSettings** se aplican solo a complementos de panel de tares y contenido. Para complementos de correo, puede especificar las ubicaciones predeterminadas de los archivos de origen y otros parámetros predeterminados en el elemento [FormSettings](../../reference/manifest/formsettings.md).

