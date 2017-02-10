# <a name="extensionpoint-element"></a>Elemento ExtensionPoint

 Define dónde expone su función un complemento en la interfaz de usuario de Office. El elemento **ExtensionPoint** es un elemento secundario de [DesktopFormFactor](./desktopformfactor.md) o de [MobileFormFactor](./mobileformfactor.md). 

## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Sí  | El tipo de punto de extensión que se está definiendo.|


## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Puntos de extensión para comandos de complemento de Word, Excel, PowerPoint y OneNote

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
|**OfficeTab**|Es obligatorio si quiere extender una pestaña de cinta de Office predeterminada (mediante **PrimaryCommandSurface**). Si usa el elemento **OfficeTab**, no puede usar el elemento **CustomTab**. Para obtener información detallada, consulte [OfficeTab](officetab.md).|
|**OfficeMenu**|Es obligatorio si agrega comandos de complemento a un menú contextual predeterminado (con **ContextMenu**). El atributo **id** debe establecerse en: <br/> - **ContextMenuText** para Excel o Word. Muestra el elemento en el menú contextual cuando se selecciona un texto y luego el usuario hace clic en él con el botón derecho. <br/> - **ContextMenuCell** para Excel. Muestra el elemento en el menú contextual cuando el usuario hace clic con el botón derecho en una celda de la hoja de cálculo.|
|**Group**|Un grupo de puntos de extensión de interfaz de usuario en una pestaña. Un grupo puede tener hasta seis controles. El atributo  **id** es obligatorio. Es una cadena con un máximo de 125 caracteres.|
|**Label**|Obligatorio. La etiqueta del grupo. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **String**. El elemento  **String** es un elemento secundario del elemento **ShortStrings**, que a su vez lo es de  **Resources**.|
|**Icon**|Obligatorio. Especifica el icono del grupo que se usará en dispositivos de factor de forma pequeños o cuando se muestren demasiados botones. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **Image**. El elemento  **Image** es un elemento secundario de **Images**, que a su vez lo es de  **Resources**. El atributo **size** determina el tamaño de la imagen en píxeles. Se necesitan tres tamaños de imagen: 16, 32 y 80. También se admiten cinco tamaños opcionales: 20, 24, 40, 48 y 64.|
|**Tooltip**|Opcional. La información sobre herramientas del grupo. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **String**. El elemento  **String** es un elemento secundario del elemento **LongStrings**, que a su vez lo es de  **Resources**.|
|**Control**|Cada grupo necesita al menos un control. Un elemento  **Control** puede ser un **Button** o un **Menu**. Use  **Menu** para especificar una lista desplegable de controles de botón. Actualmente, solo se admiten botones y menús.Consulte las secciones [Controles de botones](#button-controls) y [Controles de menú](#menu-controls) para obtener más información.<br/>**Nota** Para que sea más fácil solucionar los problemas, le recomendamos que agregue un elemento **Control** y los elementos secundarios **Resources** relacionados de uno en uno.

## <a name="extension-points-for-outlook-add-in-commands"></a>Puntos de extensión para comandos de complemento de Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (solo puede usarse en el elemento [DesktopFormFactor](./desktopformfactor.md)).
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface
Este punto de extensión coloca botones en la superficie del comando de la vista de lectura de correo. En el escritorio de Outlook aparece en la cinta.

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

#### <a name="officetab-example"></a>Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface
Este punto de extensión coloca botones en la cinta para complementos mediante el formulario de redacción de correo. 

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

#### <a name="officetab-example"></a>Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Ejemplo de CustomTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

Este punto de extensión coloca botones en la cinta para el formulario que se muestra al organizador de la reunión. 

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

#### <a name="officetab-example"></a>Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

Este punto de extensión coloca botones en la cinta para el formulario que se muestra al asistente de la reunión. 

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

#### <a name="officetab-example"></a>Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Módulo

Este punto de extensión coloca botones en la cinta para la extensión de módulo. 

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](./customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface
Este punto de extensión coloca botones en la superficie del comando de la vista de lectura de correo en el factor de forma móvil.

> **Nota:** Este tipo de elemento solo se admite en Outlook para iOS.

**Elementos secundarios**

|  Elemento |  Descripción  |
|:-----|:-----|
|  [Group](./group.md) |  Agrega un grupo de botones a la superficie de comando.  |
|  [Control](./control.md) |  Agrega un único botón a la superficie de comando.  |

Los elementos **ExtensionPoint** de este tipo solo pueden tener un elemento secundario, un elemento **Group** o un elemento **Control**.

Los elementos **Control** que se incluyen en este punto de extensión deben tener el atributo **xsi:type** establecido en `MobileButton`.

#### <a name="group-example"></a>Ejemplo de grupo
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
    <!-- one or more Control elements -->
  </Group>
</ExtensionPoint>
```

#### <a name="control-example"></a>Ejemplo de control
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Control id="mobileButton1" xsi:type="MobileButton">
    <!-- Control definition -->
  </Control>
</ExtensionPoint>
```