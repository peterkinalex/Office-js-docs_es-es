# Elemento ExtensionPoint

 Define dónde expone su funcionalidad un complemento en la interfaz de usuario de Office. El elemento **ExtensionPoint** es un elemento secundario de [FormFactor](./formfactor.md). 

## Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Sí  | El tipo de punto de extensión que se está definiendo.|


## Puntos de extensión para comandos de complemento de Word, Excel, PowerPoint y OneNote

- **PrimaryCommandSurface**: la cinta de opciones en Office.
- **ContextMenu**: el menú contextual que aparece cuando se hace clic con el botón derecho en la interfaz de usuario de Office.

Los ejemplos siguientes muestran cómo usar el elemento  **ExtensionPoint** con los valores de atributo **PrimaryCommandSurface** y **ContextMenu**, así como los elementos secundarios que hay que usar con cada uno de ellos.


 >**Importante**  En el caso de los elementos que contienen un atributo ID, asegúrese de proporcionar un identificador único. Le recomendamos que use el nombre de la compañía con su identificador. Puede seguir el siguiente formato.<CustomTab id="mycompanyname.mygroupname">


```XML
 <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Contoso Tab">
            <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
             <!-- <OfficeTab id="TabData"> -->
              <Label resid="residLabel4" />
              <Group id="Group1Id12">
                <Label resid="residLabel4" />
                <Icon>
                  <bt:Image size="16" resid="icon1_32x32" />
                  <bt:Image size="32" resid="icon1_32x32" />
                  <bt:Image size="80" resid="icon1_32x32" />
                </Icon>
                <Tooltip resid="residToolTip" />
                <Control xsi:type="Button" id="Button1Id1">

                   <!-- information about the control -->
                </Control>
                <!-- other controls, as needed -->
              </Group>
            </CustomTab>
          </ExtensionPoint>

        <ExtensionPoint xsi:type="ContextMenu">
          <OfficeMenu id="ContextMenuCell">
            <Control xsi:type="Menu" id="ContextMenu2">
                   <!-- information about the control -->
            </Control>
           <!-- other controls, as needed -->
          </OfficeMenu>
         </ExtensionPoint>
```

**Elementos secundarios**
 
|**Elemento**|**Descripción**|
|:-----|:-----|
|**CustomTab**|Es obligatorio si quiere agregar una pestaña personalizada a la cinta de opciones (con  **PrimaryCommandSurface**). Si usa el elemento  **CustomTab**, no puede usar el elemento  **OfficeTab**. El atributo  **id** es obligatorio.|
|**OfficeTab**|Es necesario si quiere extender una ficha de cinta de Office predeterminada (mediante **PrimaryCommandSurface**). Si usa el elemento **OfficeTab**, no puede usar el elemento **CustomTab**. Para obtener información detallada, consulte [OfficeTab](officetab.md).|
|**OfficeMenu**|Necesario si agrega comandos de complemento a un menú contextual predeterminado (con **ContextMenu**). El atributo **id ** debe establecerse en: <br/> - **ContextMenuText** para Excel o Word. Muestra el elemento en el menú contextual cuando se selecciona un texto y luego el usuario hace clic en él con el botón derecho. <br/> - **ContextMenuCell** para Excel. Muestra el elemento en el menú contextual cuando el usuario hace clic con el botón derecho en una celda de la hoja de cálculo.|
|**Group**|Un grupo de puntos de extensión de interfaz de usuario en una pestaña. Un grupo puede tener hasta seis controles. El atributo  **id** es obligatorio. Es una cadena con un máximo de 125 caracteres.|
|**Label**|Obligatorio. La etiqueta del grupo. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **String**. El elemento  **String** es un elemento secundario del elemento **ShortStrings**, que a su vez lo es de  **Resources**.|
|**Icono**|Obligatorio. Especifica el icono del grupo que se usará en dispositivos de factor de forma pequeños o cuando se muestren demasiados botones. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **Image**. El elemento  **Image** es un elemento secundario de **Images**, que a su vez lo es de  **Resources**. El atributo **size** determina el tamaño de la imagen en píxeles. Se necesitan tres tamaños de imagen: 16, 32 y 80. También se admiten cinco tamaños opcionales: 20, 24, 40, 48 y 64.|
|**Información sobre herramientas**|Opcional. La información sobre herramientas del grupo. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **String**. El elemento  **String** es un elemento secundario del elemento **LongStrings**, que a su vez lo es de  **Resources**.|
|**Control**|Cada grupo necesita al menos un control. Un elemento  **Control** puede ser un **Button** o un **Menu**. Use  **Menu** para especificar una lista desplegable de controles de botón. Actualmente, solo se admiten botones y menús.Consulte las secciones [Controles de botones](#controles-de-botones) y [Controles de menú](#controles-de-menú) para obtener más información.<br/>**Nota** Para que sea más fácil solucionar los problemas, le recomendamos que agregue un elemento **Control** y los elementos secundarios **Resources** relacionados de uno en uno.

## Puntos de extensión para comandos de complemento de Outlook

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (solo puede usarse en el [DesktopFormFactor](./formfactor.md).)

### CustomPane

El punto de extensión CustomPane define un complemento que se activa cuando se cumplen las reglas especificadas. Es solo para el formulario de lectura y se muestra en un panel horizontal. 

**Elementos secundarios**

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **RequestedHeight** | No |  La altura solicitada, en píxeles, para el panel de visualización cuando se ejecuta en un equipo de escritorio. Puede ser de 32 a 450 píxeles.  |
|  **SourceLocation**  | Sí |  La dirección URL del archivo de código fuente del complemento. Hace referencia a un elemento **Url** en el elemento [Resources](./resources.md).  |
|  **Rule**  | Sí |  La regla o colección de reglas que especifican cuándo se activa el complemento. Para obtener más información, consulte [Reglas de activación para complementos de Outlook](../../outlook/manifests/activation-rules.md). |
|  **DisableEntityHighlighting**  | No |  Especifica si es necesario desactivar el resaltado de entidades. |


#### Ejemplo de CustomPane
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```

### MessageReadCommandSurface
Este punto de extensión coloca botones en la superficie del comando de la vista de lectura de correo. En el escritorio de Outlook aparece en la cinta.

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la ficha de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la ficha de la cinta personalizada.  |

#### Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### MessageComposeCommandSurface
Este punto de extensión coloca botones en la cinta para complementos mediante el formulario de redacción de correo. 

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la ficha de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la ficha de la cinta personalizada.  |

#### Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Ejemplo de CustomTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### AppointmentOrganizerCommandSurface

Este punto de extensión coloca botones en la cinta para el formulario que se muestra al organizador de la reunión. 

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la ficha de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la ficha de la cinta personalizada.  |

#### Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

Este punto de extensión coloca botones en la cinta para el formulario que se muestra al asistente de la reunión. 

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la ficha de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la ficha de la cinta personalizada.  |

#### Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### Módulo

Este punto de extensión coloca botones en la cinta para la extensión de módulo. 

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la ficha de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la ficha de la cinta personalizada.  |

