
# Labs.IMessageHandler

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Interfaz que le permite definir controladores de eventos.

```
interface IMessageHandler(origin: Window, data: any, callback: Labs.Core.ILabCallback<any>): void
```


## 

 **Parámetros**


|||
|:-----|:-----|
| `origin`|La ventana del laboratorio desde la que se ha originado el mensaje.|
| `data`|Los contenidos del mensaje.|
| `callback`|Función de devolución de llamada que se desencadena una vez que se recibe el mensaje.|
