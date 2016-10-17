

# <a name="settings.remove-method"></a>Método Settings.remove
Elimina la configuración especificada.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.1|

```js
Office.context.document.settings.remove(name);
```


## <a name="parameters"></a>Parámetros



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **string**

&nbsp;&nbsp;&nbsp;&nbsp;Nombre, con distinción de mayúsculas y minúsculas, de la configuración que se debe eliminar.

    



## <a name="remarks"></a>Observaciones

 **null** es un valor válido para una configuración. Por lo tanto, si se asigna **null** a la configuración, no se eliminará del contenedor de propiedades de la configuración.


 >**Importante**: sea consciente de que el método **Settings.remove** afecta solo a la copia en memoria del contenedor de propiedades de configuración. Para continuar quitando la configuración especificada en el documento, en algún momento después de llamar al método **Settings.remove** y antes de que se cierre el complemento, debe llamar al método [Settings.saveAsync](../../reference/shared/settings.saveasync.md).


## <a name="example"></a>Ejemplo




```js
function removeMySetting() {
    Office.context.document.settings.remove('mySetting');
}
```




## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Configuración|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para crear configuraciones personalizadas en los complementos de contenido para Access.|
|1.0|Agregado|
