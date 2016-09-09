
# Labs.connect (overload)

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Inicializa una conexión con el host.

```
function connect(labHost: Core.ILabHost, callback: Core.ILabCallback<Core.IConnectionResponse>)
```


## Parámetros


|||
|:-----|:-----|
| _labHost_|Opcional. La instancia [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) a la que conectarse. Si el host no está especificado, se construirá uno mediante [Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md).|
| _callback_|Devolución de llamada que se desencadena una vez que la conexión se ha establecido.|

## Valor devuelto

Devuelve una conexión al host.

