# <a name="workbook-object-javascript-api-for-excel"></a>Objeto Workbook (API de JavaScript para Excel)

Workbook es el objeto de nivel superior que contiene los objetos de libro relacionados, como hojas de cálculo, tablas, intervalos, etc.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|application|[Aplicación](application.md)|Representa una instancia de aplicación de Excel que contiene este libro. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|bindings|[BindingCollection](bindingcollection.md)|Representa una colección de enlaces que forman parte del libro. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|functions|[Functions](functions.md)|Representa una instancia de aplicación de Excel que contiene este libro. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|Representa una colección de elementos con nombre en el ámbito del libro (intervalos y constantes con nombre). Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Representa una colección de tablas dinámicas asociadas con el libro. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|settings|[SettingCollection](settingcollection.md)|Representa una colección de configuraciones asociadas con el libro. Solo lectura.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Representa una colección de tablas asociadas con el libro. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|worksheets|[WorksheetCollection](worksheetcollection.md)|Representa una colección de hojas de cálculo asociadas con el libro. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Obtiene el intervalo seleccionado actualmente en el libro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getselectedrange"></a>getSelectedRange()
Obtiene el intervalo seleccionado actualmente en el libro.

#### <a name="syntax"></a>Sintaxis
```js
workbookObject.getSelectedRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```