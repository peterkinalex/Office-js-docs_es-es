
# LabsJS.Labs

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

El módulo LabsJS.Labs contiene el conjunto de las API clave de JavaScript que se pueden usar para crear los complementos de Office (los laboratorios). Las API proporcionan el punto de entrada para el desarrollo de laboratorios.

## Módulo de API LabsJS.Labs

El módulo Laboratorios contiene los siguientes tipos:


### Variables


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md)|Use este objeto para construir una instancia predeterminada de [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md).|

### Funciones


|||
|:-----|:-----|
|[Labs.Connect](../../reference/office-mix/labs.connect.md)|Inicializa una conexión con el host.|
|[Labs.connect (overload)](../../reference/office-mix/labs.connect-overload.md)|Inicializa una conexión con el host y proporciona parámetros de entrada.|
|[Labs.isConnected](../../reference/office-mix/labs.isconnected.md)|Inicializa una conexión con el host.|
|[Labs.getConnectionInfo](../../reference/office-mix/labs.getconnectioninfo.md)|Recupera información de configuración asociada con una conexión especificada.|
|[Labs.disconnect](../../reference/office-mix/labs.disconnect.md)|Desconecta el laboratorio del host y proporciona el estado de finalización del laboratorio.|
|[Labs.editLab](../../reference/office-mix/labs.editlab.md)|Abre el laboratorio especificado para su edición. Puede especificar los datos de configuración del laboratorio mientras esté en el modo de edición. Sin embargo, no puede editar un laboratorio mientras se esté llevando a cabo (es decir, mientras se ejecute).|
|[Labs.takeLab](../../reference/office-mix/labs.takelab.md)|Ejecuta el laboratorio especificado y habilita el envío de los resultados del laboratorio al servidor. Tenga en cuenta que no se puede ejecutar un laboratorio mientras se está editando.|
|[Labs.on](../../reference/office-mix/labs.on.md)|Agrega un nuevo controlador para un evento especificado.|
|[Labs.off](../../reference/office-mix/labs.off.md)|Quita un controlador de eventos para un evento especificado.|
|[Labs.getTimeline](../../reference/office-mix/labs.gettimeline.md)|Recupera una instancia del objeto [Labs.Timeline](../../reference/office-mix/labs.timeline.md) que se puede usar para manejar el control de reproductor de host.|
|[Labs.registerDeserializer](../../reference/office-mix/labs.registerdeserializer.md)|Deserializa un objeto JSON especificado en un objeto. Solo deben usarlo los autores de componentes.|

### Clases


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](../../reference/office-mix/labs.componentinstancebase.md)|Clase base para la inicialización de instancias del componente.|
|[Labs.ComponentInstance](../../reference/office-mix/labs.componentinstance.md)|Representa una instancia de un componente, que es una creación de instancia de un componente dado para un usuario en tiempo de ejecución. El objeto contiene una vista traducida del componente para una ejecución específica de un laboratorio.|
|[Labs.Command](../../reference/office-mix/labs.command.md)|Comando general que se usa para transmitir mensajes entre el cliente y el host.|
|[Labs.LabEditor](../../reference/office-mix/labs.labeditor.md)|El objeto **LabEditor** le permite editar un laboratorio determinado, así como obtener y establecer los datos de configuración asociados al laboratorio.|
|[Labs.LabInstance](../../reference/office-mix/labs.labinstance.md)|Una instancia de un laboratorio que está configurado para el usuario actual. Use este objeto para grabar y recuperar datos de laboratorio para el usuario.|
|[Labs.Timeline](../../reference/office-mix/labs.timeline.md)|Proporciona acceso a la característica de escala de tiempo labs.js.|
|[Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)|Un objeto contenedor que contiene y realiza un seguimiento de los valores de un laboratorio especificado. El valor puede almacenarse localmente o en el servidor.|

### Interfaces


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](../../reference/office-mix/labs.getactionscommanddata.md)|Permite recuperar los datos asociados con un comando [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md).|
|[Labs.IMessageHandler](../../reference/office-mix/labs.imessagehandler.md)|Interfaz que le permite definir controladores de eventos.|
|[Labs.ITimelineNextMessage](../../reference/office-mix/labs.itimelinenextmessage.md)|Proporciona medios para interactuar con el objeto [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx).|
|[Labs.SendMessageCommandData](../../reference/office-mix/labs.sendmessagecommanddata.md)|Datos asociados a un comando [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx).|
|[Labs.TakeActionCommandData](../../reference/office-mix/labs.takeactioncommanddata.md)|Datos asociados con un comando de realizar una acción.|

### Enumeraciones


|||
|:-----|:-----|
|[Labs.ConnectionState](../../reference/office-mix/labs.connectionstate.md)|Enumera los posibles estados de conexión del laboratorio al host.|
|[Labs.ProblemState](../../reference/office-mix/labs.problemstate.md)|Valores de estado para un laboratorio determinado.|
