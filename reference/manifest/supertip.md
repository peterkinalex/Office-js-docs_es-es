## <a name="supertip"></a>Supertip
Define una información sobre herramientas enriquecida (título y descripción). Lo usan los controles [Button](./button.md) y [Menu](./menu-control.md). 

## <a name="child-elements"></a>Elementos secundarios
|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Title](#title)        | Sí |   El texto de la sugerencia.         |
|  [Description](#description)  | Sí |  La descripción de la sugerencia.    |

## <a name="title"></a>Título
Obligatorio. Texto para la sugerencia falsa. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento [ShortStrings](./resources.md#shortstrings) del elemento [Resources](./resources.md).

## <a name="description"></a>Descripción
Obligatorio. Descripción para la sugerencia falsa. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento [LongStrings](./resources.md#longstrings) del elemento [Resources](./resources.md).

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```