
# Elemento HighResolutionIconUrl
Especifica la dirección URL de la imagen que se usa para representar su complemento de Office en la UX de inserción y la Tienda Office en pantallas con valores altos de PPP. 

 **Tipo de complemento:** Contenido, panel de tareas, correo


## Sintaxis:


```XML
<HighResolutionIconUrl DefaultValue="string " />
```


## Puede contener:

[Override](../../reference/manifest/override.md)


## Atributos



|**Atributo**|**Tipo**|**Necesario**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultValue|cadena (URL)|necesario|Especifica el valor predeterminado de esta opción, expresado para la configuración regional especificada en el elemento [DefaultLocale](../../reference/manifest/defaultlocale.md).|

## Observaciones

Para un complemento de correo, se muestra el icono en la interfaz de usuario **Archivo**  >  **Administrar complementos**. Para un complemento de contenido o panel de tareas, se muestra el icono en la interfaz de usuario **Insertar**  >  **Complementos**.

La imagen debe estar en uno de los siguientes formatos de archivo con una resolución recomendada de 64 x 64 píxeles: GIF, JPG, PNG, EXIF, BMP o TIFF. Para obtener más información, consulte la sección _Crear una identidad visual consistente para la aplicación_ en [Crear aplicaciones y complementos de la Tienda Office eficaces](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx).

