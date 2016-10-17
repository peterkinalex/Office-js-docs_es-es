

# <a name="settings.set-method"></a>Método Settings.set
Define o crea la configuración especificada.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.1|

```js
Office.context.document.settings.set(name, value);
```


## <a name="parameters"></a>Parámetros



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **string**

&nbsp;&nbsp;&nbsp;&nbsp;Nombre, con distinción entre mayúsculas y minúsculas, de la configuración que se debe establecer o crear.

    
_value_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipe: **string**, **number**, **boolean**, **null**, **object** o **array**

&nbsp;&nbsp;&nbsp;&nbsp;Especifica el valor que se debe almacenar.
    

## <a name="remarks"></a>Comentarios

El método **set** crea una nueva configuración del nombre que se especifica si no existe en el momento o establece para este una configuración ya existente en la copia en memoria del contenedor de propiedades de configuración. Tras llamar al método [Settings.saveAsync](../../reference/shared/settings.saveasync.md), el valor se almacena en el documento como la representación JSON en serie del tipo de datos correspondiente. El espacio máximo disponible para la configuración de cada complemento es de 2 MB.


 >**Importante**: sea consciente de que el método **Settings.set** afecta no solo a la copia en memoria del contenedor de propiedades de configuración. Para asegurarse de que las adiciones o los cambios en la configuración estarán disponibles en el complemento la próxima vez que se abra el documento, en algún momento después de llamar el método **Settings.set** y antes de que el complemento se cierre, debe llamar al método **Settings.saveAsync** para mantener la configuración en el documento.


## <a name="example"></a>Ejemplo




```js
function setMySetting() {
    Office.context.document.settings.set('mySetting', 'mySetting value');
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
|**Word**|v|v|v|

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
|1.1|Se ha agregado compatibilidad para las configuraciones personalizadas en complementos de contenido para Access.|
|1.0|Agregado|
