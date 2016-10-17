
# <a name="labs.valueholder"></a>Labs.ValueHolder

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Un objeto contenedor que contiene y realiza el seguimiento de valores para un laboratorio especificado. El valor puede almacenarse localmente o en el servidor.

```
class ValueHolder<T>
```


## <a name="variables"></a>Variables


|||
|:-----|:-----|
| `public var isHint: boolean`|**True** si el valor es una sugerencia.|
| `public var hasBeenRequested: boolean`|**True** si el laboratorio ha solicitado el valor.|
| `public var hasValue: boolean`|**True** si el contenedor de valor tiene actualmente el valor deseado.|
| `public var value: T`|El valor que se almacena en el contenedor.|
| `public var id: string`|El identificador del valor.|

## <a name="methods"></a>Métodos




### <a name="getvalue"></a>getValue

 `public function getValue(callback: Labs.Core.ILabCallback<T>): void`

Recupera el valor especificado.

 **Parámetros**


|||
|:-----|:-----|
| _callback_|Función de devolución de llamada que devuelve un valor especificado.|

### <a name="providevalue"></a>provideValue

 `public function provideValue(value: T): void`

Método interno que proporciona el valor al contenedor de valores.

 **Parámetros**


|||
|:-----|:-----|
| _value_|El valor para proporcionar al contenedor de valores.|
