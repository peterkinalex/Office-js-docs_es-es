
# <a name="coerciontype-enumeration"></a>Enumeración CoercionType
Especifica cómo convertir los datos que el método invocado ha devuelto o definido.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Última modificación en Buzón**|1.1|

```js
Office.CoercionType
```

## <a name="members"></a>Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.CoercionType.Html|"html"|Devuelve o establece datos como HTML.<br/><br/> **Nota**  Solo se aplica a los datos de los complementos para Word y los complementos de Outlook para Outlook (modo de redacción).|
|Office.CoercionType.Matrix|"matrix"|Devuelve o establece datos como datos tabulares sin ningún encabezado. Los datos se devuelven o se establecen como una matriz de matrices que contiene series unidimensionales de caracteres. Por ejemplo, tres filas de valores **string** en dos columnas sería: ` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`.<br/><br/> **Nota**  Solo se aplica a los datos de Excel y Word.|
|Office.CoercionType.Ooxml|"ooxml"|Devuelve o establece los datos como Office Open XML.<br/><br/> **Nota**  Solo se aplica a los datos en Word.|
|Office.CoercionType.SlideRange|"slideRange"|Devuelve un objeto JSON que contiene una matriz de identificadores, títulos e índices de las diapositivas seleccionadas. Por ejemplo, `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` para una selección de dos diapositivas.<br/><br/> **Nota**  Solo se aplica a los datos en PowerPoint al llamar al método [Document.getSelectedData](../../reference/shared/document.getselecteddataasync.md) para obtener la diapositiva actual o el rango seleccionado de diapositivas.|
|Office.CoercionType.Table|"table"|Devuelve o establece datos como datos tabulares con encabezados opcionales. Datos devueltos o establecidos como una matriz de matrices con encabezados opcionales.<br/><br/> **Nota**  Solo se aplica a los datos de Access, Excel y Word.|
|Office.CoercionType.Text|"text"|Devuelve o establece los datos como texto (**string**). Los datos se devuelven o se establecen como una serie unidimensional de caracteres.|
|Office.CoercionType.Image|"image"|Los datos se devuelven o establecen como una secuencia de imagen.<br/><br/> **Nota**  Solo se aplica a los datos de Excel, Word y PowerPoint.|
PowerPoint solo admite **Office.CoercionType.Text**,  **Office.CoercionType.Image** y **Office.CoercionType.SlideRange**.

Project solo admite **Office.CoercionType.Text**.


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|**OWA para dispositivos**|**Office para Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|v|||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, Outlook (modo de redacción) y panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para los complementos para Access.|
|1.1|Se ha agregado compatibilidad para [los complementos de Outlook con modo de redacción](../../docs/outlook/compose-scenario.md).|
|1.0|Agregado|
