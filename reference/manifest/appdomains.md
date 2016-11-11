
# <a name="appdomains-element"></a>Elemento AppDomains
Enumera los dominios además del dominio especificado en el elemento [SourceLocation](../../reference/manifest/sourcelocation.md) que el complemento de Office utilizará para cargar páginas. Para cada dominio adicional, especifique un elemento [AppDomain](../../reference/manifest/appdomain.md).

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax"></a>Sintaxis:


```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```


## <a name="contained-in"></a>Forma parte de:

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain"></a>Puede contener:

[AppDomain](../../reference/manifest/appdomain.md)


## <a name="remarks"></a>Comentarios

De forma predeterminada, el complemento puede cargar cualquier página que esté en el mismo dominio que la ubicación especificada en el elemento [SourceLocation](../../reference/manifest/sourcelocation.md). Para cargar páginas que no estén en el mismo dominio que el complemento, especifique los dominios usando los elementos **AppDomains** y **AppDomain**. Este elemento no puede estar vacío. 

Para más información, vea [Manifiesto XML de complementos de Office](../../docs/overview/add-in-manifests.md).

