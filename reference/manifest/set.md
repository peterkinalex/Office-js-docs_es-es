
# <a name="set-element"></a>Elemento Set
Especifica un conjunto de requisitos de la API de JavaScript para Office que su complemento de Office necesita para activarse.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax:"></a>Sintaxis:


```XML
<Set Name="string " MinVersion="n .n ">
```


## <a name="contained-in:"></a>Forma parte de:

[Conjuntos](../../reference/manifest/sets.md)


## <a name="attributes"></a>Atributos



|**Atributo**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Nombre|string|necesario|El nombre de un [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).|
|MinVersion|string|opcional|Especifica la versión mínima del conjunto de API que necesita el complemento. Reemplaza el valor **DefaultMinVersion** si se especifica en el elemento primario [Sets](../../reference/manifest/sets.md).|

## <a name="remarks"></a>Observaciones

Para obtener más información sobre los conjuntos de requisitos, consulte [Especificar los requisitos de la API y del host de Office](../../docs/overview/specify-office-hosts-and-api-requirements.md#specify-office-hosts-and-api-requirements).

Para obtener más información sobre el atributo **MinVersion** del elemento **Set** y del atributo **DefaultMinVersion** del elemento **Sets**, consulte [Especificar los requisitos de la API y el host de Office](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).


 >**Importante** Para los complementos de correo, solo habrá un conjunto de requisitos de `"Mailbox"` disponible. Este conjunto de requisitos contiene todo el subconjunto de API admitidas en complementos de correo para Outlook, y deberá especificar el conjunto de requisitos de `"Mailbox"` en el manifiesto del complemento de correo (no es opcional, como en el caso de los complementos de contenido y panel de tareas). Tampoco podrá declarar compatibilidad con métodos específicos en los complementos de correo.

