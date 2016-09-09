
# Elemento AllowSnapshot
Especifica si una imagen de instantánea de su complemento de contenido debe guardarse con el documento host.

 **Tipo de complemento:** Contenido


## Sintaxis:


```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```


## Forma parte de:

[OfficeApp](../../reference/manifest/officeapp.md)


## Observaciones


 **Nota de seguridad:**   **AllowSnapshot** es **true** de forma predeterminada. Esto hace que una imagen del complemento sea visible para los usuarios que abren el documento en una versión de la aplicación host que no admite complementos de Office, o proporciona una imagen estática del complemento si la aplicación host no puede conectarse al servidor que aloja el complemento. Sin embargo, esto también significa que se puede tener acceso a información confidencial que se muestra en el complemento directamente desde el documento que lo hospeda.

