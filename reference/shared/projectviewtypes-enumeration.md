
# <a name="projectviewtypes-enumeration"></a>Enumeración ProjectViewTypes
Especifica los tipos de vistas que puede reconocer el método **[ getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)**.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Agregado en**|1.0|

```
ProjectViewTypes={
    Gantt           : 1, 
    NetworkDiagram  : 2, 
    TaskDiagram     : 3, 
    TaskForm        : 4, 
    TaskSheet       : 5, 
    ResourceForm    : 6, 
    ResourceSheet   : 7, 
    ResourceGraph   : 8, 
    TeamPlanner     : 9, 
    TaskDetails     : 10, 
    TaskNameForm    : 11, 
    ResourceNames   : 12, 
    Calendar        : 13, 
    TaskUsage       : 14, 
    ResourceUsage   : 15, 
    Timeline        : 16
}
```


## <a name="members"></a>Miembros


****


|**Miembro**|**Descripción**|
|:-----|:-----|
|**Gantt**|Vista Diagrama de Gantt.|
|**NetworkDiagram**|Vista Diagrama de red.|
|**TaskDiagram**|Vista Diagrama de tareas.|
|**TaskForm**|Vista Formulario de tareas.|
|**TaskSheet**|Vista Hoja de tareas.|
|**ResourceForm**|Vista Formulario de recursos.|
|**ResourceSheet**|Vista Hoja de recursos.|
|**ResourceForm**|Vista Formulario de recursos.|
|**ResourceGraph**|Vista Gráfico de recursos.|
|**TeamPlanner**|Vista Organizador de equipo.|
|**TaskDetails**|Vista Detalles de tarea.|
|**TaskNameForm**|Vista Formulario Nombre de tarea.|
|**ResourceNames**|Vista Nombres de los recursos.|
|**Calendar**|Vista Calendario.|
|**TaskUsage**|Vista Uso de tareas.|
|**ResourceUsage**|Vista Uso de recursos.|
|**Timeline**|Vista Escala de tiempo.|

## <a name="remarks"></a>Comentarios

El método **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** devuelve el nombre y el valor de la constante **ProjectViewTypes** correspondientes a la vista activa.


## <a name="support-details"></a>Detalles de compatibilidad


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

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.0|Agregado|

## <a name="see-also"></a>Vea también



#### <a name="other-resources"></a>Otros recursos


[Método getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)
