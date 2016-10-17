
# <a name="requirements-element"></a>Elemento Requirements
Especifica el conjunto mínimo de requisitos de la API de JavaScript para Office ([conjuntos de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_sets) o métodos) que su complemento de Office necesita para activarse.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax:"></a>Sintaxis:


```XML
<Requirements>
   ...
</Requirements>
```


## <a name="contained-in:"></a>Forma parte de:

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain:"></a>Puede contener:



|**Elemento**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Conjuntos](../../reference/manifest/sets.md)|x|x|x|
|[Métodos](../../reference/manifest/methods.md)|x||x|

## <a name="remarks"></a>Observaciones

Para obtener más información sobre los conjuntos de requisitos, consulte [Especificar los requisitos de la API y del host de Office](../../docs/overview/specify-office-hosts-and-api-requirements.md).

