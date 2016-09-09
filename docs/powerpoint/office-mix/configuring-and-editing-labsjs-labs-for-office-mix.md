
# Configuración y edición de laboratorios de LabsJS para Office Mix



Aplicaciones para Office Mix proporciona métodos de office.js para obtener y establecer configuraciones de laboratorio. La configuración indica a Aplicaciones para Office Mix el tipo de laboratorio que se está creando, así como el tipo de datos que el laboratorio devolverá. Esta información se usa para recopilar y visualizar análisis.

## Obtener el editor de laboratorio

El editor de laboratorio, el objeto [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md), permite editar el laboratorio, así como obtener y establecer la configuración del laboratorio. Cuando haya terminado de editar el laboratorio, debe llamar al método  **Done**. Sin embargo, no es necesario llamar al método  **Done** excepto cuando intenta realizar o ejecutar un laboratorio que está editando. Tenga en cuenta que solo se puede abrir una instancia del laboratorio a la vez.

El siguiente código muestra cómo obtener el editor de laboratorio.




```js
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

Use los métodos  **getConfiguration** y **setConfiguration** en [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md) para almacenar la configuración de un laboratorio determinado. La configuración ([Labs.Core.IConfiguration](../../../reference/office-mix/labs.core.iconfiguration.md)) indica a Aplicaciones para Office Mix qué datos recopilará y procesará el laboratorio. Una configuración contiene información general sobre un laboratorio, como el nombre, la versión y otras opciones de configuración. La parte más importante de la configuración es la definición de los componentes de laboratorio.

El siguiente código muestra cómo establecer y obtener una configuración. Para establecer una configuración, simplemente cree el objeto de configuración y después llame al método  **setConfiguration**. Para después recuperar la configuración, llame al método  **getConfiguration** en el objeto del editor de laboratorio.




```js

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```js

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## Cerrar el editor

Para cerrar el editor, llame al método  **Done** en el editor cuando haya terminado de editar el laboratorio. Tenga en cuenta que no puede llevar a cabo y editar un laboratorio. Sin embargo, después de llamar al método **Done**, podrá editar o ejecutar el laboratorio.


## Interactuar con un laboratorio

Después de establecer la configuración del laboratorio, está listo para empezar a interactuar con él. Cuando el laboratorio se ejecuta dentro de PowerPoint, se simulan las interacciones. Sin embargo, cuando se ejecuta dentro del reproductor de lecciones de Aplicaciones para Office Mix, los datos se almacenan en la base de datos de Aplicaciones para Office Mix y se usan en análisis.


### Obtener la instancia de laboratorio

Interactúe con el laboratorio mediante el objeto [Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md), que es una instancia del laboratorio configurado del usuario actual. Para ejecutar (o "realizar") el laboratorio, llame a la función [Labs.takeLab](../../../reference/office-mix/labs.takelab.md).


```js
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

El objeto de la instancia contiene una matriz de instancias de componentes ([Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md), [Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)) que se asignan a los componentes especificados en la configuración. De hecho, una instancia es simplemente una versión transformada de la configuración que se usa para adjuntar identificadores del servidor a los objetos de instancia, así como para ocultar al usuario determinados campos cuando sea necesario (por ejemplo, sugerencias, respuestas, etc.).


### Administración del estado

El estado es un almacenamiento temporal asociado a un usuario que ejecuta un laboratorio determinado. Puede usar el almacén para conservar información entre invocaciones sucesivas de laboratorio. Por ejemplo, un laboratorio de programación podría almacenar el trabajo actual en curso del usuario.

Para  **set** el estado, use el siguiente código.




```js
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

Para  **get** el estado, use el siguiente código.




```js
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## Resultados e instancias de componentes

A continuación, se presenta una introducción de cómo implementar las instancias de los cuatro tipos de componentes, así como ejemplos breves de los métodos de componente. 

Sin embargo, en primer lugar debe familiarizarse con dos conceptos básicos cuando trabajar con instancias de componentes. El primero de ellos es el concepto de  **intentos** y **valores**.

 **Intentos**

Un intento es un intento de un usuario para completar una instancia de componente. Por ejemplo, en el caso de una pregunta de tipo test, un intento inicia cuando el usuario comienza a trabajar en el problema y termina cuando se asigna una puntuación final. A continuación, el análisis de Aplicaciones para Office Mix agrega los resultados del usuario para el problema.


 >**Nota:** Los intentos se pueden usar para todos los tipos de componente, excepto para el tipo **DynamicComponent**.

Puede recuperar los resultados de todos los intentos asociadas con una instancia de componente determinada mediante el método  **getAttempts**. Después de recuperar los resultados, el usuario puede volver a probar uno de los intentos existentes con el método  **resume** o crear un nuevo intento con el método **createAttempt**. En el ejemplo siguiente se muestra el proceso.




```js
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **Valores**

Las instancias de componentes contienen un diccionario de claves que se asignan a una matriz de valores. Puede usar la matriz para almacenar sugerencias, comentarios o cualquier otro conjunto de valores que desee asociar con el componente. La instancia de componente proporciona acceso a estos valores mediante el método  **getValues**.

Consultar un valor de sugerencia, por ejemplo, hace que el análisis marque que el usuario tomó una sugerencia. Se realiza un seguimiento de los valores por cada intento.

En el ejemplo de código siguiente se muestra cómo consultar una sugerencia.




```js
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### ActivityComponentInstance


Use el objeto  **ActivityComponentInstace** para realizar el seguimiento de la interacción de un usuario con un componente de actividad. Esta clase proporciona un método **complete** para indicar que el usuario ha terminado de interactuar con la actividad. El método puede indicar que el usuario completó una tarea asignada, terminó una lectura o cualquier otro punto final asociado con la actividad. El código siguiente muestra cómo usar el método **complete**.


```js
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### ChoiceComponentInstance


Use el objeto  **ChoiceComponentInstance** para realizar el seguimiento de la interacción de un usuario con un componente de elección. Los componentes de elección son problemas que presentan al usuario una lista de opciones para seleccionar. Puede o no haber una respuesta correcta. La clase proporciona dos métodos principales: **getSubmissions** y **submit**. El método  **getSubmissions** permite recuperar envíos almacenados anteriormente; el método **submit** permite que se almacene un nuevo envío. Los siguientes ejemplos de código ilustran el uso de los métodos.


```js
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```js
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### InputComponentInstance


Use el objeto  **InputComponentInstance** para realizar el seguimiento de la interacción de un usuario con un componente de entrada. La clase proporciona dos métodos principales: **getSubmission** y **submit**. El método  **getSubmissions** permite recuperar los envíos almacenados previamente; el método **submit** permite almacenar un nuevo envío. El siguiente fragmento de código ilustra cómo usar el método **getSubmissions**.


```js
var submissions = this._attempt.getSubmissions();
```

Al usar el método  **submit**, tenga en cuenta que el objeto  **InputComponentAnswer** representa la respuesta enviada y el objeto **InputComponentResult** contiene el resultado. El valor devuelto es un objeto **InputComponentSubmission** que contiene la respuesta, el resultado y una marca de tiempo que indica cuándo se envió el resultado.




```js
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### DynamicComponentInstance


Use el objeto  **DynamicComponentInstance** para realizar el seguimiento de la interacción de un usuario con un componente dinámico. Los métodos principales en esta clase son **getComponents**,  **createComponent** y **close**.

El método  **getComponents** permite recuperar una lista de instancias de componente creadas anteriormente, tal como se muestra en el ejemplo siguiente.




```js
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

El método  **createComponent** crea un nuevo componente y devuelve esa instancia de componente, tal como se muestra en el ejemplo siguiente.




```js
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

Use el método  **close** para indicar que ha terminado de usar el componente dinámico para crear nuevos componentes. Tenga en cuenta que también puede usar un método booleano **isClosed** para comprobar si se ha cerrado la instancia de componente dinámico. El código siguiente muestra cómo usar el método **close**.




```js
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## Recursos adicionales



- [Complementos de Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Tutorial: Crear su primer laboratorio para Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
