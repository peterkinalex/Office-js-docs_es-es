
# Labs.takeLab

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Ejecuta el laboratorio especificado y habilita el envío de los resultados del laboratorio al servidor. Tenga en cuenta que no se puede ejecutar un laboratorio mientras se está editando.

```
function takeLab(callback: Core.ILabCallback<LabInstance>): void
```


## Parámetros


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|El método de devolución de llamada se desencadena una vez que se ha creado el objeto [Labs.LabInstance](../../reference/office-mix/labs.labinstance.md).|
