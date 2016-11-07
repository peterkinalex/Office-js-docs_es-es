
# <a name="javascript-api-for-office-reference"></a>Referencia de la API de JavaScript para Office

La API de JavaScript para Office le permite crear aplicaciones web que interactúen con los modelos de objetos de las aplicaciones host de Office. Su aplicación hará referencia a la biblioteca office.js, que es un cargador de scripts. La biblioteca office.js carga los modelos de objetos que se aplican a la aplicación de Office que realiza la ejecución del complemento. Puede usar los siguientes modelos de objetos de JavaScript:


1. Comunes (obligatorios): interfaces de programación de aplicaciones (API) que se introdujeron con Office 2013. Este modelo de objetos se carga para **todas las aplicaciones host de Office** y conecta la aplicación de complemento con la aplicación del cliente de Office. El modelo de objetos contiene interfaces de programación de aplicaciones (API) específicas de los clientes de Office e interfaces de programación de aplicaciones que se aplican a varias aplicaciones host de los clientes de Office. Todo el contenido de la secciones **API compartida** y **Outlook** corresponde a las interfaces de programación de aplicaciones (API) comunes. El espacio de nombres **Microsoft.Office.WebExtension** (al que, de forma predeterminada, se hace referencia con el alias [Office](../reference/shared/office.md) en el código) contiene objetos que puede usar para escribir scripts que interactúen con el contenido de los documentos, las hojas de cálculo, las presentaciones, los elementos de correo y los proyectos de Office desde sus complementos de Office. Deberá usar estas interfaces de programación de aplicaciones (API) comunes si el complemento tiene como destino Office 2013 o una versión posterior. Este modelo de objetos usa devoluciones de llamada.

1. Específicos del host: interfaces de programación de aplicaciones (API) que se introdujeron con **Office 2016**. Este modelo de objetos proporciona objetos específicos del host fuertemente tipados  que corresponden a los conocidos objetos que se ven cuando se usan clientes de Office y que representan el futuro de las interfaces de programación de aplicaciones (API) de JavaScript para Office. Actualmente, las interfaces de programación de aplicaciones (API) específicas del host incluyen la [API de JavaScript para Word](../reference/word/word-add-ins-reference-overview.md) y la [API de JavaScript para Excel](../reference/excel/application.md). Este modelo de objetos usa compromisos.

Seleccione el cliente de Office en la lista desplegable situada encima de la tabla de contenido para filtrar el contenido basado en la aplicación host de destino.

## <a name="supported-host-applications"></a>Aplicaciones host compatibles
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word

Obtenga más información acerca de los [hosts compatibles y otros requisitos](../docs/overview/requirements-for-running-office-add-ins.md).

## <a name="open-api-specifications"></a>Especificaciones de la API pública

Cuando diseñemos y desarrollemos nuevas interfaces de programación de aplicaciones (API) para complementos de Office, estarán disponibles para que pueda enviar sus comentarios en la página [Especificaciones de la API abierta](openspec.md). Descubra qué nuevas características están en proceso y envíe sus comentarios acerca de nuestras especificaciones de diseño.

