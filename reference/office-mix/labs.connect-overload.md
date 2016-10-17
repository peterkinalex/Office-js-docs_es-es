
# <a name="labs.connect-(overload)"></a>Labs.connect (sobrecarga)

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Inicializa una conexión con el host.

```
function connect(labHost: Core.ILabHost, callback: Core.ILabCallback<Core.IConnectionResponse>)
```


## <a name="parameters"></a>Parámetros


|||
|:-----|:-----|
| _labHost_|Opcional. La instancia [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) a la que conectarse. Si el host no está especificado, se construirá uno mediante [Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md).|
| _callback_|Devolución de llamada que se desencadena una vez que la conexión se ha establecido.|

## <a name="return-value"></a>Valor devuelto

Devuelve una conexión al host.

