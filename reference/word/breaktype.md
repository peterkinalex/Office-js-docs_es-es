# BreakType (API de JavaScript para Word)

Especifica la forma de un salto.

_Se aplica a: Word 2016, Word para iPad, Word para Mac_

A continuación se incluyen los tipos de salto compatibles en la API.

| **Valor**         | **Tipo** | **Descripción**     |
|:-----------------|:--------|:----|
|column| | Salto de columna en el punto de inserción. |
|line| | Salto de línea. |
|lineClearLeft| | Salto de línea. |
|lineClearRight| | Salto de línea. |
|next| | Salto de sección en la página siguiente. |
|page| | Salto de página en el punto de inserción.|
|sectionContinuous| | Nueva sección sin un correspondiente salto de página.|
|sectionEven| string | Salto de sección con la siguiente sección empezando en la siguiente página par. Si el salto de sección se encuentra en una página impar, Word deja en blanco la siguiente página impar.|
|sectionOdd| string | Salto de sección con la siguiente sección empezando en la siguiente página impar. Si el salto de sección se encuentra en una página impar, Word deja en blanco la siguiente página par.|
|textWrapping| string | Finaliza la línea actual y obliga a que el texto continúe debajo de una imagen, tabla u otro elemento. El texto continúa en la siguiente línea en blanco que no contiene ninguna tabla alineada con el margen izquierdo o derecho.|

## Detalles de compatibilidad
Use el [conjunto de requisitos](../office-add-in-requirement-sets.md) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).