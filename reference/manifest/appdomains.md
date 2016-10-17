
# <a name="appdomains-element"></a>Elemento AppDomains
Especifica los dominios adicionales que el complemento de Office usará para cargar páginas.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax:"></a>Sintaxis:


```XML
<AppDomains>
   ...
</AppDomains>
```


## <a name="contained-in:"></a>Forma parte de:

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain:"></a>Puede contener:

[AppDomain](../../reference/manifest/appdomain.md)


## <a name="remarks"></a>Comentarios

Los elementos **AppDomains** y **AppDomain** se usan para especificar los dominios adicionales que son distintos al que se especifica en el elemento [SourceLocation](../../reference/manifest/sourcelocation.md). Para obtener más información, consulte el [manifiesto XML de los complementos de Office](../../docs/overview/add-in-manifests.md).

