
# Enumeración ValueFormat
Especifica si se debe aplicar su formato correspondiente a los valores que devuelve el método que se ha invocado (por ejemplo, números y fechas).

|||
|:-----|:-----|
|**Hosts:**|Excel, Project y Word|
|**Agregado en**|1,0|

```
Office.ValueFormat
```


## Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.ValueFormat.Formatted|"formatted"|Devuelve datos con formato.|
|Office.ValueFormat.Unformatted|"unformatted"|Devuelve datos sin formato.|

## Comentarios

Por ejemplo, si el parámetro _valueFormat_ se especifica como `"formatted"`, los números con formato de moneda o las fechas con el formato dd/mm/aa de la aplicación host conservarán su formato. Sin embargo, si el parámetro _valueFormat_ se especifica como `"unformatted"`, se devolverán las fechas con la forma de sus respectivos números de serie secuenciales subyacentes.


## Detalles de compatibilidad


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
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
