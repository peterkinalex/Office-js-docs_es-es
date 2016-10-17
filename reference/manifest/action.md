# <a name="action-element"></a>Elemento Action
 Especifica la acción que se realiza cuando el usuario selecciona los controles de [Botón](./button-control.md) o [Menú](./menu-control.md).
 
## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Sí  | Tipo de acción que se va a realizar|


## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Especifica el nombre de la función que se va a ejecutar. |
|  [SourceLocation](#sourcelocation) |    Especifica la ubicación del archivo de origen para esta acción. |
  

## <a name="xsi:type"></a>xsi:type
Este atributo especifica el tipo de acción que se realiza cuando el usuario selecciona el botón. Puede ser uno de las siguientes:
- ExecuteFunction
- ShowTaskpane

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
