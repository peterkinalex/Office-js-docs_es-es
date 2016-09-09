
# Instrucciones para crear laboratorios para Office Mix con LabsJS



La biblioteca de LabsJS (labs.js) admite escritura de Complementos de Office especializados (denominados laboratorios) que se integran con Aplicaciones para Office Mix. A continuación, Aplicaciones para Office Mix representa los laboratorios mediante Microsoft PowerPoint. Mientras llamamos a estos componentes "laboratorios", vamos a aclarar que lo que estamos creando son Complementos de Office especiales que son Office Mix Add-ins.

El contenido de LabsJS ayuda a implementar la API de JavaScript labs.js mediante consejos y ejemplos. Esta biblioteca se basa en la [API de JavaScript para Office](../../../reference/javascript-api-for-office.md) (Office.js) y proporciona una capa de abstracción que está optimizada para add-ins insertados en Aplicaciones para Office Mix.


## Instrucciones generales


Las siguientes son algunas instrucciones generales que le ayudarán a la hora de escribir add-ins mediante la API LabJS.


### Scripts

Dado que la biblioteca de labs.js es una capa de abstracción de office.js y, por lo tanto, tiene una dependencia en office.js, los archivos de biblioteca office.js y labs.js deben incluirse en los proyectos de desarrollo. 

Use la siguiente dirección para hacer referencia a la biblioteca de office.js: `<script src="https://sforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>`.

La biblioteca de labs.js se incluye en el SDK de LabsJS. Como alternativa, puede hacer referencia a la biblioteca labs.js en una CDN en <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. Tenga en cuenta que la versión de producción del laboratorio tiene que hacer referencia a la versión almacenada en la CDN.


 >**Nota:** Además del archivo de JavaScript (labs-1.0.4.js), proporcionamos un archivo de definición TypeScript de la API de laboratorios (labs-1.0.4.d.ts). El archivo de definición se creó con TypeScript versión 0.9.1.1.


### Devoluciones de llamada y tratamiento de errores

Varios métodos de la API labs.js funcionan de forma asincrónica. Para esas operaciones, la API adopta una interfaz estándar de devolución de llamadas,  **ILabCallback**. 


```js
function(err, result) {
}
```

El método de devolución de llamada toma dos parámetros,  _err_ y _result_. El campo  _err_ permanece **null** a menos que haya un error. El campo _result_ devuelve el resultado de la operación.

La operación de devolución de llamada nunca se desencadena inmediatamente, incluso si el resultado está disponible de inmediato. En su lugar, se desencadena en una ejecución separada del bucle de evento de JavaScript (por medio de la llamada  **setTimeout**). Mediante la adopción de esta definición de devolución de llamada, se puede integrar fácilmente labs.js con la API de la promesa de su elección. Por ejemplo, puede sustituir promesas jQuery para estas devoluciones de llamada con un método simple traducción, tal como se muestra en el ejemplo siguiente.




```js
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### Host de laboratorio y DefaultLabHost

El host de laboratorio ( **ILabHost**) es el controlador subyacente que admite el desarrollo de laboratorios. De forma predeterminada, se establece en un host que se integra con office.js.

Para fines de prueba, y para ejecutar el laboratorio en labhost.html, debe cambiar a un host que funcione en el entorno de simulación. En el siguiente ejemplo de código se muestra cómo hacer esto con un parámetro de consulta. Como alternativa, puede cambiar  **DefaultHostBuilder** para integrar el complemento de laboratorio con una plataforma diferente.




```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### Inicialización

La inicialización establece la ruta de comunicación entre el laboratorio y el host. Para inicializar el laboratorio, llame a lo siguiente.


```js
Labs.connect((err, connectionResponse) => {});
```

Después de inicializar, puede llamar a otros métodos de la API labs.js. El parámetro  _connectionResponse_ contiene información sobre el host, el usuario y otra información relacionada con la conexión. Para obtener más información sobre los valores devueltos, consulte [Labs.Core.IConnectionResponse](../../../reference/office-mix/labs.core.iconnectionresponse.md).


### Formato de hora

Labs.js almacena los números como milisegundos transcurridos desde el 1 de enero de 1970, hora UTC. Esto coincide con el formato de fecha del JavaScript [objeto Date](http://msdn.microsoft.com/en-us/library/ie/cd9w2te4%28v=vs.94%29.aspx),


### Escala de tiempo

El laboratorio también puede interactuar con la escala de tiempo del reproductor lecciones. La escala de tiempo permite que el laboratorio indique al reproductor de lecciones para que avance a la siguiente diapositiva. El objeto de escala de tiempo se recupera mediante una llamada al método  **Labs.getTimeline**.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## Administración de eventos


La API de eventos LabsJS realiza un seguimiento de los eventos específicos de laboratorio y permite agregar controladores de eventos para que pueda responder a los eventos o actuar acorde con estos. Los métodos de eventos, de los cuales hay tres, están en el objeto  **EventTypes**:  **ModeChanged**,  **Activate** y **Deactivate**. 


### Cambio de modo

El evento  **ModeChanged** se desencadena cuando el laboratorio especificado cambia del modo de edición al modo de vista. El modo de edición está visible cuando el laboratorio se ve en modo de edición de PowerPoint. El modo de vista está visible cuando PowerPoint muestra la presentación con diapositivas o cuando el laboratorio se muestra en el reproductor de lecciones de Aplicaciones para Office Mix. El modo de vista siempre debería mostrar lo que el usuario ve al realizar el laboratorio. El modo de edición permite al usuario configurar el laboratorio.

Los datos en el objeto  **ModeChangedEventData** que se pasan a la devolución de llamada contienen información sobre el modo actual. El código siguiente muestra cómo usar el evento **ModeChanged**.




```js
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### Activar

El evento  **activate** se desencadena cuando la diapositiva de PowerPoint en la que se encuentra el laboratorio se vuelve activa en el reproductor de lecciones.


```js
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### Desactivar

El evento  **deactivate** se desencadena cuando la diapositiva de PowerPoint en la que se encuentra el laboratorio ya no es la diapositiva activa.


```js
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### Escala de tiempo

El laboratorio también puede interactuar con la escala de tiempo del reproductor lecciones. La escala de tiempo permite que el laboratorio indique al reproductor de lecciones para que avance a la siguiente diapositiva. El objeto de escala de tiempo se recupera mediante una llamada al método  **Labs.getTimeline**.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## Recursos adicionales



- [Complementos de Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
