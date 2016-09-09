
# Labs.registerDeserializer

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Deserializa un objeto JSON especificado en un objeto. Solo deben usarlo los autores de componentes.

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## Parámetros


|**Nombre**|**Descripción**|
|:-----|:-----|
|json|La instancia [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) para deserializar.|

## Valor devuelto

Devuelve una instancia [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md).

