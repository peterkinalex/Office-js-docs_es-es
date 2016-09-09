#Objeto UI.Dialog
El objeto que se devuelve cuando se llama al método [displayDialogAsync](officeui.displaydialogasync.md).

## Miembros
| Miembro	       | Tipo   |Descripción|
|:---------------|:--------|:----------|
|close|Función|Permite que el complemento cierre el cuadro de diálogo.|
|addEventHandler|Función|Registra un controlador de eventos. Los dos eventos compatibles son: <ul><li>DialogMessageReceived. Se desencadena cuando el cuadro de diálogo envía un mensaje a su elemento principal.</li><li>DialogEventReceived. Se desencadena cuando el cuadro de diálogo se ha cerrado o descargado.</li></ul> |


### close()
Llamado desde una página principal para cerrar el cuadro de diálogo correspondiente.     
```js    
[dialogObject].close();    
``` 

#### Parámetros    
Ninguno 

#### Valores devueltos    
void  


#### Ejemplos
Para obtener ejemplos, vea el tema [Método DisplayDialogAsync](officeui.displaydialogasync.md).
