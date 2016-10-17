
# <a name="gototype-enumeration"></a>Enumeración GoToType
Especifica el tipo de lugar u objeto hacia el que se debe navegar.

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint y Word|
|**Agregado en**|1.1|

```js
Office.GoToType
```


## <a name="members"></a>Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|**Clientes compatibles**|
|:-----|:-----|:-----|:-----|
|Office.GoToType.Binding|"binding"|Va a un objeto de enlace mediante el ID de enlace especificado.|Excel</br>Word|
|Office.GoToType.NamedItem|"namedItem"|Se dirige a un elemento con el nombre de dicho elemento (por ejemplo, el nombre asignado a una tabla o rango). En Excel, puede usar cualquier referencia estructurada para una tabla o un rango con nombre: "Worksheet2!Table1"|Excel|
|Office.GoToType.Slide|"slide"|Va a una diapositiva utilizando el Id. especificado.|PowerPoint|
|Office.GoToType.Index|"index"|Va al índice especificado por número de diapositiva o enumeración:</br>**Office.Index.First**</br>**Office.Index.Last**</br>**Office.Index.Next**</br>**Office.Index.Previous**|PowerPoint|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.


Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Agregado|
