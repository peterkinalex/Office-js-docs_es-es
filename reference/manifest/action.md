# <a name="action-element"></a>Elemento Action
 Especifica la acción que se realiza cuando el usuario selecciona los controles de [Botón](./control.md#button-control) o [Menú](./control.md#menu-dropdown-button-controls).
 
## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Sí  | Tipo de acción que se va a realizar|


## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Especifica el nombre de la función que se va a ejecutar. |
|  [SourceLocation](#sourcelocation) |    Especifica la ubicación del archivo de origen para esta acción. |
|  [TaskpaneId](#taskpaneid) | Especifica el ID del contenedor del panel de tareas.|
|  [SupportsPinning](#supportspinning) | Especifica que un panel de tareas admite el anclado, lo que provoca que el panel de tareas siga abierto aunque el usuario cambie la selección.|
  

## <a name="xsitype"></a>xsi:type
Este atributo especifica el tipo de acción que se realiza cuando el usuario selecciona el botón. Puede ser uno de las siguientes:

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

Elemento obligatorio cuando **xsi:type** es "ExecuteFunction". Especifica el nombre de la función que se va a ejecutar. La función está incluida en el archivo especificado en el elemento [FunctionFile](./functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation
Elemento obligatorio cuando  **xsi:type** es "ShowTaskpane". Especifica la ubicación del archivo de origen para esta acción. El atributo **resid** debe establecerse en el valor del atributo **id** de un elemento **Url** en el elemento [Urls](./resources.md#urls) del elemento [Resources](./resources.md).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId
Elemento opcional cuando **xsi: Type** es "ShowTaskpane". Especifica el identificador del contenedor de panel de tareas. Cuando haya varias acciones de "ShowTaskpane", utilice una **TaskpaneId** diferente si desea un panel independiente para cada uno. Utilice el mismo **TaskpaneId** para distintas acciones que comparten el mismo panel. Cuando los usuarios eligen comandos que comparten la misma **TaskpaneId**, el contenedor del panel permanecerá abierto pero el contenido del panel se reemplazará por la correspondiente acción "SourceLocation". 

>**Nota:** Este elemento no es compatible con Outlook.

El ejemplo siguiente muestra dos acciones que comparten el mismo **TaskpaneId**. 


```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

## <a name="supportspinning"></a>SupportsPinning
Elemento opcional cuando **xsi: Type** es "ShowTaskpane". Los elementos que contengan [VersionOverrides](./versionoverrides.md) deben tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`. Incluya este elemento con el valor `true` para admitir el anclado de paneles de tareas. El usuario podrá "anclar" el panel de tareas, lo provocará que permanezca abierto cuando se cambie la selección.

>**Nota:** Actualmente, este elemento solo se admite en Outlook 2016.

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```