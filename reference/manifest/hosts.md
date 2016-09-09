# Elemento Hosts

Especifica la aplicación cliente de Office donde se activará el complemento de Office. Contiene una colección de elementos **Host** y su configuración. 

Cuando se incluye en el nodo [VersionOverrides](./versionoverrides.md), este elemento reemplaza el elemento **Hosts** en la parte principal del manifiesto. 

## Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Host](#host)    |  Sí   |  Describe un host y su configuración. |

> ** Nota: ** Outlook requiere que `Hosts` contenga una definición `Host` para `MailHost`.

---- 

## Elemento Host
Especifica un tipo de aplicación de Office individual donde debe activarse el complemento (como "documento", "libro", "presentación", "proyecto", "buzón" y "bloc de notas").

### Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Sí  | Describe el host de Office al que se aplica esta configuración.|

### Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  Sí   |  Define el factor de forma afectado. |


### xsi:type
Controla a qué host de Office (Word, Excel, PowerPoint, Outlook, OneNote) se aplica también la configuración contenida. El valor debe ser uno de los siguientes:

- `MailHost` (Outlook)    


## Ejemplo de hosts 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
