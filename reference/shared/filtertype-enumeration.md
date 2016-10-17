
# <a name="filtertype-enumeration"></a>Enumeración FilterType
Especifica si se debe aplicar el filtrado desde la aplicación host al recuperar los datos.

|||
|:-----|:-----|
|**Hosts:**|Excel, Project y Word|
|**Modificado por última vez en**|1.1|

```js
Office.FilterType
```


## <a name="members"></a>Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.FilterType.All|"all"|Devolver todos los datos (sin filtrado de la aplicación host).|
|Office.FilterType.OnlyVisible|"onlyVisible"|Devolver solo los datos visibles (con filtrado de la aplicación host).|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.


Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Project**|v|||
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad

|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
