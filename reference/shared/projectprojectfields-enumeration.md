
# Enumeración ProjectProjectFields
Especifica los campos del proyecto que están disponibles como parámetro del método **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)**.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Agregado en**|1,0|

```
ProjectProjectFields={
    CurrencyDigits: 0, 
    CurrencySymbol: 1, 
    CurrencySymbolPosition: 2, 
    DurationUnits: 3,
    GUID: 4, 
    Finish: 5, 
    Start: 6, 
    ReadOnly: 7, 
    VERSION: 8, 
    WorkUnits: 9, 
    ProjectServerUrl: 10, 
    WSSUrl: 11, 
    WSSList: 12
}
```


## Miembros


****


|**Miembro	**|**Descripción**|
|:-----|:-----|
|**CurrencyDigits**|El número de dígitos después del decimal para la moneda.|
|**CurrencySymbol**|El símbolo de moneda.|
|**CurrencySymbolPosition**|La ubicación del símbolo de moneda: No especificada = -1; Antes del valor, sin espacio ($0) = 0; Después del valor, sin espacio (0$) = 1; Antes del valor, con un espacio ($ 0) = 2; Después del valor, con un espacio (0 $) = 3.|
|**GUID**|El GUID del proyecto.|
|**Finish**|La fecha de finalización del proyecto.|
|**Iniciar**|La fecha de inicio del proyecto.|
|**ReadOnly**|Especifica si el proyecto es de solo lectura.|
|**VERSIÓN**|La versión del proyecto.|
|**WorkUnits**|Las unidades de trabajo del proyecto (por ejemplo, horas o días).|
|**ProjectServerUrl**|La dirección URL de Project Web App para los proyectos que se almacenan en Project Server.|
|**WSSUrl**|La dirección URL de SharePoint para los proyectos sincronizados con una lista de SharePoint.|
|**WSSList**|El nombre de la lista de SharePoint para los proyectos sincronizados con una lista de tareas.|

## Comentarios

Se puede usar una constante **ProjectProjectFields** como parámetro del método **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)**.


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1,0|Agregado|

## Vea también



#### Otros recursos


[Método getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)
