# Elemento VersionOverrides

Elemento raíz que contiene la información de los comandos del complemento implementados por el complemento. **VersionOverrides** es un elemento secundario del elemento [OfficeApp](./officeapp.md) del manifiesto. Este elemento se admite en la versión 1.1 del esquema del manifiesto y posterior, pero se define en el esquema de la versión 1.0 de VersionOverrides. 

## Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **xmlns**       |  Sí  |  Ubicación del esquema, que tiene que ser `http://schemas.microsoft.com/office/mailappversionoverrides`.|
|  **xsi:type**  |  Sí  | Versión del esquema. En este momento, el único valor válido es `VersionOverridesV1_0`. |


## Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **Descripción**    |  No   |  Describe el complemento. Esto reemplaza el elemento `Description` en cualquier parte principal del manifiesto. El texto de la descripción está contenido en un elemento secundario del elemento **LongString**, contenido en el elemento [Resources](./resources.md). El atributo `resid` del elemento **Description** está establecido en el valor del atributo `id` del elemento `String` que contiene el texto.|
|  **Requisitos**  |  No   |  Especifica el conjunto de requisitos mínimos y la versión de Office.js que necesita el complemento. Esto reemplaza el elemento `Requirements` en cualquier parte principal del manifiesto.| 
|  [Hosts](./hosts.md)                |  Sí  |  Especifica una colección de hosts de Office. El elemento Hosts secundario reemplaza el elemento Hosts en cualquier parte principal del manifiesto.  |
|  [Resources](./resources.md)    |  Sí  | Define una colección de recursos (cadenas, direcciones URL e imágenes) a las que hacen referencia otros elementos del manifiesto.|



### Ejemplo de VersionOverrides
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information on resources -->
   </Resources>
</VersionOverrides>
...
</OfficeApp>
```
