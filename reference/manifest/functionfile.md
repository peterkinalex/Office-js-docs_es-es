# <a name="functionfile-element"></a>Elemento FunctionFile

Especifica el archivo de código fuente para las operaciones que expone un complemento a través de comandos que ejecutan una función de JavaScript en lugar de mostrar la UI. El elemento **FunctionFile** es un elemento secundario de [FormFactor](./formfactor). El atributo **resid** del elemento **FunctionFile** está establecido en el valor del atributo **id** de un elemento **Url** en el elemento **Resources** que contiene la dirección URL a un archivo HTML que contiene o carga todas las funciones de JavaScript que usan los botones de comandos del complemento sin interfaz de usuario, como define el [Control element](control.md).

El siguiente es un ejemplo del elemento **FunctionFile**.


```XML
<DesktopFormFactor>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!-- information about this extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
```

El código de JavaScript que hay en el archivo HTML indicado en el elemento **FunctionFile** tiene que llamar a `Office.initialize` y definir funciones con nombre que toman un solo parámetro: `event`. Las funciones tienen que usar la API [item.notificationMessages](../../../reference/outlook/Office.context.mailbox.item.md) para indicar progreso, éxito o error al usuario. También tiene que llamar a [event.completed](../../../reference/shared/event.completed.md) al finalizar la ejecución. El nombre de las funciones se usa en el elemento **FunctionName** para los botones sin interfaz de usuario.

A continuación se muestra un ejemplo de un archivo HTML que define una función **trackMessage**.

```js
Office.intialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

El código siguiente muestra cómo implementar la función que se usa en  **FunctionName**.




```js
        // The initialize function must be run each time a new page is loaded.
        (function () {
            Office.initialize = function (reason) {
               // If you need to initialize something you can do so here.
            };
        })();

            // Your function must be in the global namespace.
        function writeText(event) {

            // Implement your custom code here. The following code is a simple example.

            Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === "failed") {
                        // Show error message.
                    }
                    else {
                        // Show success message.
                    }
                });
           // Calling event.completed is required. event.completed lets the platform know that processing has completed.
       event.completed();
        }
```


 >**Importante** La llamada a **event.completed** indica que se ha controlado correctamente el evento. Cuando se llama a una función varias veces, como puede ser haciendo varios clics en el mismo comando de complemento, se ponen automáticamente en cola de todos los eventos. El primer evento se ejecuta automáticamente, mientras que los demás eventos permanecen en la cola. Cuando la función llama a **event.completed**, se ejecuta la siguiente llamada en cola a esa función. Debe implementar **event.completed**; de lo contrario, no se ejecutará la función.
