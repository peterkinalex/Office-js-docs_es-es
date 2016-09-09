## Supertip
Define una información sobre herramientas enriquecida (título y descripción). Lo usan los controles [Button](./button.md) y [Menu](./menu-control.md). 

## Elementos secundarios
|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Título](#título)        | Sí |   El texto de la sugerencia.         |
|  [Descripción](#descripción)  | Sí |  La descripción de la sugerencia.    |

## Título
Obligatorio. Texto para la sugerencia falsa. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento [ShortStrings](./resources.md#shortstrings) del elemento [Resources](./resources.md).

## Descripción
Obligatorio. Descripción para la sugerencia falsa. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento [LongStrings](./resources.md#longstrings) del elemento [Resources](./resources.md).

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```