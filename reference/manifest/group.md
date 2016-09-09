# Elemento Group
Define un grupo de puntos de extensión de UI en una ficha.  En las pestañas personalizadas, el complemento puede crear hasta 10 grupos. Cada grupo está limitado a 6 controles, independientemente de la pestaña donde aparezca. Los complementos están limitados a una pestaña personalizada.

## Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [id](#id)  |  Sí  | Un identificador único para el grupo.|

## Elementos secundarios
|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Label](#label)      | Sí |  La etiqueta de CustomTab o de un grupo.  |
|  [Control](#control)    | Sí |  Colección de uno o más objetos Control.  |

## Atributo id
Necesario. Identificador único para el grupo. Es una cadena con un máximo de 125 caracteres. Debe ser único dentro del manifiesto o el grupo no podrá procesarse.

## Label 
Obligatorio. La etiqueta del grupo. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento [ShortStrings](./resources.md#shortstrings) del elemento [Resources](./resources.md).

## Control
Un grupo necesita como mínimo un control. Actualmente, solo se admiten [botones](./control.md#button-control) y [menús](./menu.md#menu-control). 

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```