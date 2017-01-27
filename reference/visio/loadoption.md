# <a name="object-load-options-javascript-api-for-visio"></a>Objeto Load Options (API de JavaScript para Visio)

>**Nota:** Las API de JavaScript para Visio no están disponibles actualmente para su uso en entornos de producción o de versión preliminar.

Representa un objeto que se puede pasar al método Load para especificar el conjunto de propiedades y las relaciones que van a cargar tras la ejecución del método **sync()** que sincroniza los estados entre los objetos de Visio y los correspondientes objetos proxy de JavaScript. Usa opciones como los parámetros Select y Expand para especificar un conjunto de propiedades que se va a cargar en el objeto y también permite la paginación en la colección.

También se puede suministrar una cadena que contenga las propiedades de relaciones que se cargarán, o bien proporcionar una matriz que contenga la lista de propiedades y relaciones que se cargarán. Vea el ejemplo siguiente.

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## <a name="properties"></a>Propiedades

| Propiedad | Tipo  | Descripción |
|:---------|:------|:------------|
|select    |object |Proporciona una lista delimitada por comas o una matriz de nombres de parámetros/relaciones que se cargarán al realizar una llamada executeAsync, como "propiedad1, relación1", [ "propiedad1", "relación1"]. Opcional.|
|expand    |object |Proporciona una lista delimitada por comas o una matriz de nombres de relaciones que se cargarán al realizar una llamada executeAsync, como "relación1, relación2", [ "relación1", "relación2"]. Opcional.|
|top       |int    |Especifica el número de elementos de la colección consultada que se deben incluir en el resultado. Opcional.|
|skip      |int    |Especifica el número de elementos de la colección que se deben omitir y no se incluyen en el resultado. Si se especifica **top**, la selección de resultados empezará después de omitir el número especificado de elementos. Opcional.|

