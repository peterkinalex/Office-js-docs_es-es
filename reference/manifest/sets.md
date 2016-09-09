
# Elemento Sets
Especifica el subconjunto mínimo de la API de JavaScript para Office que su complemento de Office necesita para activarse.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## Sintaxis:


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## Forma parte de:

[Requisitos](../../reference/manifest/requirements.md)


## Puede contener:

[Set](../../reference/manifest/set.md)


## Atributos



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|opcional|Especifica el valor de atributo predeterminado de **MinVersion** elementos [Set](../../reference/manifest/set.md) secundarios. El valor predeterminado es "1.1".|

## Observaciones

Para obtener más información sobre los conjuntos de requisitos, consulte [Especificar los requisitos de la API y del host de Office](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Para obtener más información sobre el atributo **MinVersion** del elemento **Set**  y del atributo **DefaultMinVersion** del elemento **Sets**, consulte [Definir el elemento Requirements en el manifiesto](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

