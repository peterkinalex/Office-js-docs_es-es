
# Elemento Host
Especifica un tipo individual de aplicación de Office en el que se debe activar el complemento.

> **Importante**: La sintaxis del elemento **Host** varía dependiendo de si el elemento está definido dentro del [manifiesto básico](#basic-manifest) o dentro del nodo [VersionOverrides](#versionoverrides-node). Sin embargo, las funciones son las mismas.  


## Manifiesto básico

Cuando se define en el manifiesto básico (bajo [OfficeApp](./officeapp.md)), el tipo de host está determinado en el atributo `Name`.   

### Atributos
| Atributo     | Tipo   | Necesario | Descripción                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Nombre](#name) | string | necesario | El nombre del tipo de aplicación host de Office. |


### Nombre
Especifica el tipo de host al que se dirige este complemento. El valor debe ser uno de los siguientes:

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### Ejemplo
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

---

## Nodo VersionOverrides
Cuando se define en [VersionOverrides](./versionoverrides), el tipo de host está determinado en el atributo `xsi:type`. 

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

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## Ejemplo de host 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
