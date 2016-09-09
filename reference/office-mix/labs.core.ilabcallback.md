
# Labs.Core.ILabCallback

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

La interfaz para controlar los métodos de devolución de llamada Labs.js.

```
interface ILabCallback<T>
```


## Firma de devolución de llamada

 `(err: any, data: T): void`

 **Parámetros de devolución de llamada**


|||
|:-----|:-----|
| _err_|**Null** si no se producen errores. No **null** si se produce un error.|
| _data_|Los datos devueltos con la devolución de llamada.|
