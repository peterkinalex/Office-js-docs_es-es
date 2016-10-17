#<a name="ui.dialog-object"></a>Objeto UI.Dialog
El objeto que se devuelve cuando se llama al método [displayDialogAsync](officeui.displaydialogasync.md).

## <a name="members"></a>Miembros
| Miembro	       | Tipo   |Descripción|
|:---------------|:--------|:----------|
|close|Función|Permite que el complemento cierre el cuadro de diálogo.|
|addEventHandler|Función|Registra un controlador de eventos. Los dos eventos compatibles son: <ul><li>DialogMessageReceived. Se desencadena cuando el cuadro de diálogo envía un mensaje a su elemento principal.</li><li>DialogEventReceived. Se desencadena cuando el cuadro de diálogo se ha cerrado o descargado.</li></ul> |


### <a name="close()"></a>close()
Llamado desde una página principal para cerrar el cuadro de diálogo correspondiente.     
```js    
[dialogObject].close();    
``` 

#### <a name="parameters"></a>Parámetros    
Ninguno 

#### <a name="returns"></a>Valores devueltos    
void  


#### <a name="examples"></a>Ejemplos
Para obtener ejemplos, vea el tema [Método DisplayDialogAsync](officeui.displaydialogasync.md).
