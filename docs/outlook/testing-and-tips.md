
# Implementar e instalar complementos de Outlook para probarlos


Como parte del proceso de desarrollo del complemento de Outlook, se encontrará repetidamente implementando e instalando el complemento para probarlo, lo cual consiste en los siguientes pasos:


1. Creación de un archivo de manifiesto que describa el complemento.
    
2. Implementación de los archivos de la IU del complemento en un servidor web.
    
3. Instalación del complemento en su buzón de correo.
    
4. Pruebe el complemento, haciendo los cambios apropiados a la IU o los archivos de manifiesto, y repitiendo los pasos 2 y 3 para probar los cambios.
    

## Creación de un archivo de manifiesto para el complemento.

Cada complemento se describe en un manifiesto XML, un documento que proporciona al servidor información sobre el complemento, proporciona información descriptiva sobre el componente al usuario e identifica la ubicación del archivo HTML de la interfaz de usuario del componente. Puede almacenar el manifiesto en una carpeta local o en un servidor siempre que el servidor de Exchange del buzón de correo que está probando pueda tener acceso a dicha ubicación. Se asume que el manifiesto se guardará en una carpeta local. Para obtener información sobre cómo crear un archivo de manifiesto [Manifiestos de complementos de Outlook](../outlook/manifests/manifests.md). 


## Implementación de un complemento en un servidor web

Puede usar HTML y JavaScript para crear la interfaz de usuario del complemento. El archivo de origen generado se almacena en un servidor web al que puede tener acceso el servidor Exchange que hospeda el complemento. El archivo de origen es identificado por el elemento secundario  **SourceLocation** en el elemento [DesktopSettings](http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx), el elemento [TableSettings](http://msdn.microsoft.com/en-us/library/5c89cc7c-7ae0-49c9-fdd5-4c52118228f6%28Office.15%29.aspx) o el elemento [PhoneSettings](http://msdn.microsoft.com/en-us/library/13e4eae3-8e8c-fd55-a1c2-3297b485f327%28Office.15%29.aspx) que se especifican en el archivo de manifiesto del complemento.

Una vez que haya desplegado los archivos de interfaz de usuario para el complemento, puede actualizar la interfaz de usuario del complemento y su comportamiento reemplazando el archivo HTML almacenado en el servidor web con una nueva versión del archivo HTML.


## Instalación del complemento


Una vez que se haya realizado la preparación del archivo del manifiesto del complemento y se hayan implementado los archivos de la interfaz de usuario del complemento a un servidor web al que se pueda obtener acceso, se puede instalar el complemento para el servidor de Exchange usando un cliente enriquecido de Outlook, Outlook Web App o OWA para dispositivos; o bien, iniciando cmdlets remotos de Windows PowerShell.


### Instalación de un complemento en un cliente enriquecido de Outlook

Puede instalar un complemento si el buzón está en Exchange Online, Exchange 2013 o en una versión posterior. En Outlook para Windows, puede instalar complementos desde la vista Backstage de Office Fluent. Seleccione **Archivo** y **Administrar complementos**. Esto le permitirá iniciar sesión en el Centro de administración de Exchange. Después de iniciar sesión, continúe con el proceso de instalación (paso 4) en la sección siguiente.

En Outlook para Mac, seleccione **Administrar complementos** en el extremo derecho de la barra de complementos y, después, inicie sesión en el Centro de Administración de Exchange. Continúe con el paso 4 en la sección siguiente.


### Instalación de un complemento mediante Outlook Web App o Outlook.com

Para usar Outlook Web App (OWA) para instalar un complemento de Outlook, siga los pasos siguientes:


1. Vaya a la dirección URL de OWA de su organización o Outlook.com e inicie sesión.
    
2. Haga clic en el icono de engranaje en la esquina superior derecha y elija **Administrar complementos**.
    
3. Seleccione el signo más (**+**) para agregar un complemento nuevo.
    
4. En la lista desplegable, seleccione **Agregar desde archivo** (si guardó el manifiesto en una carpeta local).
    
5. Navegue hasta la ruta de acceso al archivo de manifiesto y seleccione **Instalar**.
    
6. Seleccione el nombre de usuario en la esquina superior derecha de la ventana y seleccione **Mi correo** para cambiar a su correo electrónico y probar el complemento.
    

>**Nota** Si no usa ninguna de las siguientes características para desarrollar el complemento: 
- Inquilino desarrollador de Office 365
- Herramientas de desarrollo de Office 365 de Napa
- Visual Studio

Y, si no tiene el rol "Mis aplicaciones personalizadas", como mínimo, en su Exchange Server, solo podrá instalar complementos desde la Tienda Office. Para probar el complemento o instalar complementos en general especificando un nombre de archivo o dirección URL del manifiesto del complemento, debe solicitar al administrador de Exchange que le proporcione los permisos necesarios.

El administrador de Exchange puede ejecutar el siguiente cmdlet de PowerShell para asignar a un solo usuario los permisos necesarios. En este ejemplo, wendyri es el alias de correo electrónico del usuario.

```New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"```

Si es necesario, el administrador puede ejecutar el cmdlet siguiente para asignar a varios usuarios permisos necesarios parecidos:

```$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}```

Para más información sobre el rol Mis aplicaciones personalizadas, vea [Rol Mis aplicaciones personalizadas](http://technet.microsoft.com/en-us/library/aa0321b3-2ec0-4694-875b-7a93d3d99089%28EXCHG.150%29.aspx). 

Al usar Office 365, Napa o Visual Studio para desarrollar complementos, se le asignará el rol de administrador de la organización, que le permite instalar complementos con un archivo o una URL en EAC, o bien con los cmdlets de PowerShell.


### Instalación de un complemento con PowerShell remoto

Una vez que haya creado una sesión remota de Windows PowerShell en su servidor de Exchange, puede instalar un complemento de Outlook con el cmdlet  **New-App** con el siguiente comando PowerShell.


```
New-App -URL:"http://<fully-qualified URL">
```

La dirección URL completa es la ubicación del archivo de manifiesto del complemento que preparó para su complemento.

Puede usar los siguientes cmdlets adicionales de PowerShell para administrar los complementos para un buzón de correo:


-  **Get-App**: enumera los complementos que están activados para un buzón de correo.
    
-  **Set-App**: activa o desactiva un complemento para un buzón de correo.
    
-  **Remove-App**: elimina del servidor Exchange un complemento previamente instalado.
    

## Recursos adicionales



- [Complementos de Outlook](../outlook/outlook-add-ins.md)
    
- [Solucionar errores de usuario con los complementos de Office](../testing/testing-and-troubleshooting.md)
    
