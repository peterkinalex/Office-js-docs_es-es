# <a name="getstarted-element"></a>Elemento GetStarted

Proporciona información usada por la llamada que se muestra cuando el complemento se instala en hosts de Word, Excel, PowerPoint y OneNote. El elemento **GetStarted** es un elemento secundario de [DesktopFormFactor](./desktopformfactor.md).

## <a name="child-elements"></a>Elementos secundarios

| Elemento                       | Obligatorio | Descripción                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Sí      | Define dónde expone su funcionalidad un complemento.     |
| [Description](#description)   | Sí      | Una dirección URL de un archivo que contiene funciones de JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | No       | Una dirección URL de una página que explica el complemento en detalle.   |


## <a name="title"></a>Título 
Necesario. El título que se usa para la parte superior de la llamada. El atributo **resid** hace referencia a un identificador válido del elemento [ShortStrings](./resources.md#shortstrings) en la sección [Recursos](./resources.md).

## <a name="description"></a>Descripción
Necesario. La descripción / contenido del cuerpo de la llamada. El atributo **resid** hace referencia a un identificador válido del elemento [LongStrings](./resources.md#longstrings) en la sección [Recursos](./resources.md).

## <a name="learnmoreurl"></a>LearnMoreUrl
Necesario. La dirección URL de una página donde el usuario puede encontrar más información sobre el complemento. El atributo **resid** hace referencia a un identificador válido del elemento [Urls](./resources.md#urls) en la sección [Recursos](./resources.md).

> **NOTA:** **LearnMoreUrl** no se representa actualmente en los clientes de Word, Excel o PowerPoint. Se recomienda agregar esta dirección URL a todos los clientes de forma que la dirección URL se represente cuando esté disponible. 
