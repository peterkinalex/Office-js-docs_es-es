
# <a name="iconurl-element"></a>Elemento IconUrl
Especifica la dirección URL de la imagen que se usa para representar su complemento de Office en la UX de inserción y la Tienda Office.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax:"></a>Sintaxis:


```XML
<IconUrl DefaultValue="string " />
```


## <a name="can-contain:"></a>Puede contener:

[Override](../../reference/manifest/override.md)


## <a name="attributes"></a>Atributos



|**Atributo**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|necesario|Especifica el valor predeterminado de esta opción, expresado para la configuración regional especificada en el elemento [DefaultLocale](../../reference/manifest/defaultlocale.md).|

## <a name="remarks"></a>Observaciones

Para un complemento de correo, se muestra el icono en la interfaz de usuario **Archivo**  >  **Administrar complementos** (Outlook) o **Configuración**  >  **Administrar complementos** (Outlook Web App). Para un complemento de contenido o panel de tareas, se muestra el icono en la interfaz de usuario **Insertar**  >  **Complementos**. Para todos los tipos de complemento, también se usa el icono en el sitio de la Tienda Office si publica el complemento en la Tienda Office.

La imagen debe estar en uno de los formatos siguientes: GIF, JPG, PNG, EXIF, BMP o TIFF. En el caso de las aplicaciones de contenido y de panel de tareas, la imagen especificada debe ser de 32 x 32 píxeles. En el caso de las aplicaciones de correo la imagen debe ser de 64 x 64 píxeles. También debería especificar un icono para usarlo con las aplicaciones host de Office que se ejecutan en pantallas con valores altos de PPP mediante el elemento [HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md). Para obtener más información, consulte la sección _Crear una identidad visual uniforme para su aplicación_ en [Crear aplicaciones y complementos de Office eficaces](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx).

