
# LabsJS.Labs.Core.Actions
Proporciona una visión general de la API de JavaScript de LabJS.Labs.Core.Actions.

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Estas API representan las operaciones de un laboratorio, indicando los comportamientos actuales de este. Las API son útiles si está creando nuevos componentes o desarrollando conexiones con un nuevo controlador (que no sea Office Mix).

## LabsJS.Labs.Core.Actions API module

El módulo Acciones contiene los siguientes tipos:


### Interfaces


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](../../reference/office-mix/labs.core.actions.iclosecomponentoptions.md)|El componente que se debe cerrar.|
|[Labs.Core.Actions.ICreateAttemptOptions](../../reference/office-mix/labs.core.actions.icreateattemptoptions.md)|El componente asociado al intento.|
|[Labs.Core.Actions.ICreateAttemptResult](../../reference/office-mix/labs.core.actions.icreateattemptresult.md)|El resultado de crear un intento para el componente determinado.|
|[Labs.Core.Actions.ICreateComponentOptions](../../reference/office-mix/labs.core.actions.icreatecomponentoptions.md)|Crea un nuevo componente.|
|[Labs.Core.Actions.ICreateComponentResult](../../reference/office-mix/labs.core.actions.icreatecomponentresult.md)|El resultado [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) de la creación de un nuevo componente.|
|[Labs.Core.Actions.IGetValueResult](../../reference/office-mix/labs.core.actions.igetvalueresult.md)|El resultado de una acción de obtención de valor.|
|[Labs.Core.Actions.ISubmitAnswerResult](../../reference/office-mix/labs.core.actions.isubmitanswerresult.md)|El resultado de enviar una respuesta para un intento.|
|[Labs.Core.Actions.IAttemptTimeoutOptions](../../reference/office-mix/labs.core.actions.iattempttimeoutoptions.md)|Opciones disponibles para la acción de tiempo de espera del intento actual.|
|[Labs.Core.Actions.IGetValueOptions](../../reference/office-mix/labs.core.actions.igetvalueoptions.md)|Opciones disponibles de la operación de obtención de valor.|
|[Labs.Core.Actions.IResumeAttemptOptions](../../reference/office-mix/labs.core.actions.iresumeattemptoptions.md)|Opciones asociadas al intento de reanudación.|
|[Labs.Core.Actions.ISubmitAnswerOptions](../../reference/office-mix/labs.core.actions.isubmitansweroptions.md)|Opciones disponibles para la acción de envío de respuestas.|

### Variables


|||
|:-----|:-----|
| `var CloseComponentAction: string`|Cierra el componente e indica que no habrá más acciones futuras contra este.|
| `var CreateAttemptAction: string`|Acción para crear un nuevo intento.|
| `var CreateComponentAction: string`|Acción para crear un nuevo componente.|
| `var AttemptTimeoutAction: string`|Intento de una acción de tiempo de espera.|
| `var GetValueAction: string`|Acción para recuperar un valor asociado a un intento.|
| `var ResumeAttemptAction: string`|Reanudar la acción de intento. Se usa para indicar que el usuario está reanudando el trabajo en un intento determinado.|
| `var SubmitAnswerAction: string`|Acción para enviar una respuesta para un intento determinado.|
