
# <a name="labs.core.ivalueinstance"></a>Labs.Core.IValueInstance

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Una instancia de objeto [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) que contiene datos de valor, si los hubiera.

```
interface IValueInstance
```


## <a name="properties"></a>Propiedades


|||
|:-----|:-----|
| `valueId: string`|Identificador del valor que representa esta instancia.|
| `isHint: boolean`|Valor booleano **true** si este valor se considera una sugerencia.|
| `hasValue: boolean`|Valor booleano **true** si la información de la instancia contiene el valor.|
| `value?: any`|El valor. Este parámetro puede o no puede establecerse dependiendo de si se ha ocultado.|
