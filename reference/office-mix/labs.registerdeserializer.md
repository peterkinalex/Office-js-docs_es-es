
# <a name="labs.registerdeserializer"></a>Labs.registerDeserializer

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Deserializa un objeto JSON especificado en un objeto. Solo lo deben usar los autores del componente.

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## <a name="parameters"></a>Parámetros


|**Nombre**|**Descripción**|
|:-----|:-----|
|json|La instancia [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) para deserializar.|

## <a name="return-value"></a>Valor devuelto

Devuelve una instancia [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md).

