# Elemento Icon
Define los elementos de **Image** para los controles de [Button](./button.md) y [Menu](./menu-control.md).

## Elementos secundarios
|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Image](#image)        | Sí |   resid de una imagen que se usará         |

## Image
Imagen del botón. El atributo **resid** tiene que establecerse en el valor del atributo **id** de un elemento **Image** en el elemento **Images** del elemento [Resources](./resources.md). El atributo **size** indica el tamaño en píxeles de la imagen. Se necesitan tres tamaños de imágenes (16, 32 y 80 píxeles), mientras que se admiten otros cinco tamaños (20, 24, 40, 48 y 64 píxeles).|


```xml
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
```  