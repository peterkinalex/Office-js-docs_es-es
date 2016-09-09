
# Labs.Timeline

 _**Hace referencia a:** aplicaciones para Office | Complementos de Office | Office Mix | PowerPoint_

Proporciona acceso a la característica de escala de tiempo labs.js.

```
class Timeline
```


## Métodos




### method

 `function constructor(labsInternal: Labs.LabsInternal)`

Crea una nueva instancia en la clase **Timeline**.


### next

 `public function next(completionStatus: Labs.Core.ICompletionStatus, callback: Labs.Core.ILabCallback<void>): void`

Indica que la escala de tiempo debería avanzar a la diapositiva siguiente.

 **Parámetros**


|||
|:-----|:-----|
| _completionStatus_|Indica el estado actual del laboratorio.|
| _callback_|Función de devolución de llamada que se desencadena cuando el laboratorio se ha movido a la diapositiva siguiente.|
