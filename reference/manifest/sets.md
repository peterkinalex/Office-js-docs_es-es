
# <a name="sets-element"></a>Elemento Sets
Especifica el subconjunto mínimo de la API de JavaScript para Office que su complemento de Office necesita para activarse.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax:"></a>Sintaxis:


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## <a name="contained-in:"></a>Forma parte de:

[Requirements](../../reference/manifest/requirements.md)


## <a name="can-contain:"></a>Puede contener:

[Conjunto](../../reference/manifest/set.md)


## <a name="attributes"></a>Atributos



|**Atributo**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|opcional|Especifica el valor de atributo predeterminado de **MinVersion** elementos [Set](../../reference/manifest/set.md) secundarios. El valor predeterminado es "1.1".|

## <a name="remarks"></a>Observaciones

Para obtener más información sobre los conjuntos de requisitos, consulte [Especificar los requisitos de la API y del host de Office](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Para obtener más información sobre el atributo **MinVersion** del elemento **Set**  y del atributo **DefaultMinVersion** del elemento **Sets**, consulte [Definir el elemento Requirements en el manifiesto](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

