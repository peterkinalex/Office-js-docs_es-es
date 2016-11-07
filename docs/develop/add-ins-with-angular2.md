# <a name="tips-for-creating-office-addins-with-angular-2"></a>Sugerencias para crear complementos de Office con Angular 2 

Este artículo proporciona instrucciones para utilizar Angular 2 para crear un complemento de Office como una aplicación de una sola página.

>**Nota:** ¿Tiene algo que aportar basándose en su experiencia con Angular 2 para crear complementos de Office? Puede contribuir a este artículo en [GitHub](https://github.com/OfficeDev/office-js-docs) o darnos su opinión enviando un [problema](https://github.com/OfficeDev/office-js-docs/issues) en el repositorio. 

Para obtener un ejemplo de complementos de Office creado con la infraestructura de Angular 2, consulte [Complemento de comprobación de estilo de Word basado en Angular 2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>La secuencia de inicio debe estar dentro de Office.initialize

En cualquier página que llame a las API de JavaScript de Office, Word o Excel, el código debe asignar un método a la propiedad `Office.initialize`. (Si no tiene ningún código de inicialización, el cuerpo del método puede ser simplemente símbolos "`{}`" vacíos, pero no debe dejar la propiedad `Office.initialize` sin definir. Para obtener más información, consulte [Inicializar el complemento](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)). Office llama a este método inmediatamente tras inicializar las bibliotecas de JavaScript de Office.

**Se debe llamar al código de secuencia de arranque Angular dentro del método que se asigna a `Office.initialize`** para asegurarse de que las bibliotecas de JavaScript de Office se han inicializado en primer lugar. El ejemplo siguiente muestra cómo hacer esto de forma sencilla. El código debe estar en el archivo main.ts del proyecto.

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
    import { AppModule } from './app.module';
    Office.initialize = function () {
        const platform = platformBrowserDynamic();
        platform.bootstrapModule(AppModule);
  };
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>Usar la estrategia de ubicación de hash en la aplicación Angular

Es posible que no pueda desplazarse de una ruta a otra si no especifica la estrategia de ubicación de hash. Puede hacerlo de una de las dos formas siguientes: Primero, puede especificar un proveedor para la estrategia de ubicación en el módulo de su aplicación, tal y como se muestra en el ejemplo siguiente. Va en el archivo app.module.ts.

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity
    @NgModule({
        providers: [
            {provide: LocationStrategy, useClass: HashLocationStrategy},
            // Other providers suppressed
        ],
        // Other module properties suppressed
  })
  export class AppModule {}
``` 

Si las rutas se definen en un módulo de enrutamiento independiente, hay una manera alternativa para especificar la estrategia de ubicación de hash. En archivo *.ts de su módulo de enrutamiento, pase un objeto de configuración a la función `forRoot` que especifica la estrategia. A continuación puede ver un ejemplo del código. 

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity
    const routes: Routes = // route definitions go here
    @NgModule({
      imports: [ RouterModule.forRoot(routes, {useHash: true}) ],
      exports: [ RouterModule ]
    })
    export class AppRoutingModule {}
```   


## <a name="consider-wrapping-fabric-components-with-angular-2-components"></a>Considere envolver los componentes Fabric con componentes Angular 2.

Recomendamos utilizar el estilo [Office UI Fabric](http://dev.office.com/fabric#/fabric-js) en sus complementos. Fabric incluye componentes que vienen con varias versiones, incluyendo una versión [basada en TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Considere el uso de componentes de Fabric en el complemento envolviéndolos en componentes Angular 2. Para ver un ejemplo que muestra cómo hacer esto, consulte [Complemento de comprobación de estilo de Word en Angular 2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Tenga en cuenta, por ejemplo, cómo el componente Angular definido en [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) importa el archivo TextField.ts, en el que se define el componente de Fabric. 


## <a name="using-the-office-dialog-api-with-angular"></a>Utilizar el cuadro de diálogo de Office API con Angular

El cuadro de diálogo del complemento de Office API le permite abrir una página en un cuadro de diálogo semimodal que puede compartir información con la página principal, que suele encontrarse en un panel de tareas. 

El método [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) toma un parámetro que especifica la dirección URL de la página que se debe abrir en el cuadro de diálogo. El complemento puede tener una página HTML independiente (diferente de la página base) para pasar a este parámetro, o puede pasar la dirección URL de una ruta en la aplicación Angular. 

Es importante recordar que si se pasa una ruta, el cuadro de diálogo crea una nueva ventana con su propio contexto de ejecución. La página base y su código de inicialización y arranque ejecutan otra vez este nuevo contexto y las variables se establecen en sus valores iniciales en el cuadro de diálogo. Por lo que esta técnica inicia una segunda instancia de la aplicación de la página en el cuadro de diálogo. El código que cambia las variables en el cuadro de diálogo no cambia la versión del panel de tareas de estas variables. De igual forma, el cuadro de diálogo tiene su propio almacenamiento de sesión, que no es accesible desde el código en el panel de tareas.  


## <a name="forcing-an-update-of-the-dom"></a>Forzar una actualización del DOM

En cualquier aplicación Angular 2, es posible que las notificaciones para actualizar el DOM en ocasiones no se activen. El marco proporciona un método `tick()` en el objeto `ApplicationRef` que forzará una actualización. A continuación puede ver un ejemplo del código.

```js
import { ApplicationRef } from '@angular/core';
    export class MyComponent {
        constructor(private appRef: ApplicationRef) {}
        myMethod() {
            // Code that changes the DOM is here
            appRef.tick();
        }
}
``` 

## <a name="using-observables"></a>Usar observables

Angular 2 utiliza RxJS (extensiones reactivas para JavaScript) y RxJS introduce objetos `Observable` y `Observer` para implementar procesamiento asíncrono. Esta sección proporciona una breve introducción al uso de `Observables`; para obtener más información, consulte la documentación oficial de [RxJS](http://reactivex.io/rxjs/).

Un `Observable` es, en cierto modo, parecido a un objeto `Promise`: lo devuelve inmediatamente una llamada asincrónica, pero no se puede resolver hasta un tiempo después. Sin embargo, mientras que un `Promise` es un valor único (que puede ser un objeto de conjunto), un `Observable` es una matriz de objetos (posiblemente con un solo miembro). Esto permite al código llamar a [métodos de matriz](http://www.w3schools.com/jsref/jsref_obj_array.asp) como `concat`, `map` y `filter` en objetos `Observable`. 

### <a name="pushing-instead-of-pulling"></a>Introducir en vez de extraer

El código "extrae" objetos `Promise` asignados a variables, pero los objetos `Observable` "introducen" sus valores en los objetos que se *suscriben* al `Observable`. Los suscriptores son objetos `Observer`. La ventaja de la arquitectura de inserción es que se pueden agregar nuevos miembros a la matriz `Observable` con el tiempo. Cuando se agrega un nuevo miembro, todos los objetos `Observer` que suscritos al `Observable` reciben una notificación. 

El `Observer` está configurado para procesar cada objeto nuevo (que recibe el nombre de objeto "siguiente") con una función. (También está configurado para responder a un error y a una notificación de finalización. Consulte la sección siguiente para obtener un ejemplo). Por este motivo, los objetos `Observable` pueden utilizarse en un mayor número de escenarios distintos que los objetos `Promise`. Por ejemplo, además de devolver un `Observable` desde una llamada de AJAX, igual que devolvería un `Promise`, un `Observable` puede devolverse desde un controlador de eventos, como el controlador del evento "modificado" para un cuadro de texto. Cada vez que un usuario escribe texto en el cuadro de diálogo, todos los objetos `Observer` suscritos reaccionan inmediatamente con el texto más reciente o el estado actual de la aplicación como entrada. 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a>Esperar hasta que se hayan completado todas las llamadas asincrónicas

Cuando desee asegurarse de que una devolución de llamada se ejecuta solo cuando todos los miembros de un conjunto de objetos `Promise` se ha resuelto, utilice el método `Promise.all()`.

```js
myPromise.all([x, y, z]).then(// TODO: Callback logic goes here.)
``` 

Para hacer lo mismo con un objeto `Observable`, se utiliza el método [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).  

```js
var source = Rx.Observable.forkJoin([x, y, z]);

var subscription = source.subscribe(
  function (x) {
    // TODO: Callback logic goes here
  },
  function (err) {
    console.log('Error: ' + err);
  },
  function () {
    console.log('Completed');
  });
``` 

