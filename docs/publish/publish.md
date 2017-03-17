
# <a name="deploy-and-publish-your-office-add-in"></a>Implementar y publicar un complemento de Office

Puede usar distintos métodos a la hora de implementar el complemento de Office para pruebas o para distribuirlo a los usuarios:

|**Método**|**Finalidad**|
|:---------|:------------|
|[Instalación de prueba](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Como parte del proceso de desarrollo para probar la ejecución del complemento en Windows, Office Online, iPad o Mac.|
|[Centro de administración de Office 365 (versión preliminar)](#office-365-admin-center-preview)|En un entorno en la nube o híbrido, para distribuir el complemento a los usuarios de su organización.|
|[Tienda Office]|Para distribuir el complemento a los usuarios de forma pública.|
|[Catálogo de SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|En un entorno local, para distribuir el complemento a los usuarios de su organización.|
|[Servidor de Exchange](#outlook-add-in-deployment)|En un entorno local o en línea, para distribuir complementos de Outlook a los usuarios.|

Las opciones disponibles dependen del tipo de complemento que cree y del host de Office al que esté destinado.

>**Nota:** Si va a publicar el complemento en la Tienda Office, asegúrese de que se ajusta a las [directivas de validación de la Tienda Office](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Por ejemplo, para superar la validación, el complemento debe funcionar en todas las plataformas que sean compatibles con los métodos especificados (para obtener más información, consulte la [sección 4.12](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably) y la [página Disponibilidad y hosts de los complementos de Office](https://dev.office.com/add-in-availability)).

## <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Opciones de implementación de complementos para Word, Excel y PowerPoint

| Punto de extensión            | Instalación de prueba | Centro de administración de Office 365 (versión preliminar) |Tienda Office| Catálogo de SharePoint*  |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| Contenido         | X           | X                  | X                               | X|
| Panel de tareas       | X           | X                  | X                               | X|
| Comando           | X           | X                  | X                               |  |

&#42; Los catálogos de SharePoint no son compatibles con Office 2016 para Mac.

## <a name="deployment-options-for-outlook-add-ins"></a>Opciones de implementación para complementos de Outlook

| Punto de extensión     | Transferencia local | Servidor Exchange | Tienda Office |
|:---------|:-----------:|:---------------:|:------------:|
| Aplicación de correo | X           | X               | X            |
| Comando  | X           | X               | X            |


Para obtener información sobre cómo pueden los usuarios finales adquirir, insertar y ejecutar complementos, consulte [Empezar a usar el complemento de Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

## <a name="office-365-admin-center-preview-deployment"></a>Implementación del Centro de administración de Office 365 (versión preliminar)

El Centro de administración de Office 365 permite a los administradores implementar fácilmente complementos de Word, Excel y PowerPoint para los usuarios o grupos de su organización. Los complementos que se implementan a través del centro de administración están disponibles de inmediato en las aplicaciones de Office de los usuarios, sin que sea necesario realizar ninguna configuración por parte del cliente. Se pueden implementar complementos internos, así como complementos proporcionados por proveedores de software independientes a través del centro de administración.

En la actualidad, el centro de administración es compatible con los siguientes escenarios:

- Implementación centralizada de complementos nuevos y actualizados para individuos, grupos o una organización.
- Compatibilidad con varias plataformas como por ejemplo Windows y Office Online, y próximamente Mac.
- Implementación en idioma inglés y los inquilinos en todo el mundo.
- Implementación de complementos hospedados en la nube.
- Instalación automática al inicio de la aplicación de Office.
- Direcciones URL de complementos hospedadas en un firewall.
- Implementación de complementos de la Tienda Office (disponible próximamente).

<!--
The admin center also includes a pre-deployment validation checking service.
-->

Los esfuerzos futuros en escenarios de implementación de complementos se enfocarán en el Centro de administración de Office 365. Por tanto, recomendamos utilizar el centro para implementar los complementos en su organización, si esta cumple los requisitos previos.

### <a name="prerequisites-for-admin-center-deployment"></a>Requisitos previos para la implementación desde el centro de administración 

Para poder implementar complementos a través del centro de administración, su organización debe cumplir los siguientes criterios:

- Los usuarios deben ejecutar Office 2016, compilación 7070 o posterior.
- Los usuarios deben iniciar sesión en Office 2016 con su cuenta profesional o educativa.
- La organización debe usar el servicio de identidad de Active Directory (AD Azure) de Azure.

El centro de administración no es compatible con lo siguiente:

- Complementos destinados a las versiones de Word, Excel o PowerPoint de Office 2013.
- Un servicio de directorio local.
- Implementación de complementos de SharePoint.
- Implementación de complementos en Office Online Server.
- Implementación de complementos COM/VSTO.

Para implementar complementos de SharePoint o complementos destinados a Office 2013, use un [catálogo de complementos de SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Importante:** Los catálogos de complementos de SharePoint no admiten características de complemento que se hayan implementado en el nodo [VersionOverrides](../../reference/manifest/versionoverrides.md) del manifiesto del complemento, como [comandos de complemento](../design/add-in-commands.md). 

Para implementar complementos COM/VSTO, use ClickOnce o Windows Installer. Para obtener más información, consulte [Implementar una solución de Office](https://msdn.microsoft.com/en-us/library/bb386179.aspx).

## <a name="sharepoint-catalog-deployment"></a>Implementación desde un catálogo de SharePoint

Un catálogo de complementos de SharePoint es una colección de sitios especial que se puede crear para hospedar complementos de Word, Excel y PowerPoint. Puesto que los catálogos de SharePoint no son compatibles con las nuevas características de los complementos implementadas en el nodo VersionOverrides del manifiesto (incluidos los comandos del complemento), le recomendamos que, si es posible, use la implementación centralizada a través del centro de administración (versión preliminar). De forma predeterminada, los comandos de los complementos que se implementan a través de un catálogo de SharePoint se abren en un panel de tareas.

Si tiene pensado implementar complementos en un entorno local, utilice un catálogo de SharePoint. Para obtener más información, vea [Publicar complementos de panel de tareas y de contenido en un catálogo de SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> **Nota:** Los catálogos de SharePoint no son compatibles con Office 2016 para Mac. Para implementar complementos de Office en clientes Mac, debe enviarlos a la [Tienda Office]. 

## <a name="outlook-add-in-deployment"></a>Implementación de complementos de Outlook

En entornos locales y en línea donde no se usa el servicio de identidad de Azure AD, los complementos de Outlook se pueden implementar mediante el servidor de Exchange. 

La implementación de complementos de Outlook requiere:

- Office 365, Exchange Online o Exchange Server 2013 o posterior
- Outlook 2013 o posterior

Para asignar los complementos a los inquilinos, se usa el Centro de administración de Exchange para cargar un manifiesto directamente desde un archivo o una dirección URL, o se agrega un complemento desde la Tienda Office. Para asignar complementos a usuarios individuales, debe usar Exchange PowerShell. Para obtener más información, consulte [Instalación o eliminación de aplicaciones para Outlook en la organización](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx) en TechNet.


## <a name="additional-resources"></a>Recursos adicionales

- [Implementar e instalar complementos de Outlook para probarlos](../outlook/testing-and-tips.md) 
- [Enviar complementos y aplicaciones web a la Tienda Office][Tienda Office]
- [Instrucciones de diseño para complementos de Office](../design/add-in-design)
- [Crear complementos de la Tienda Office efectivos](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Solucionar errores de usuario con los complementos de Office](../testing/testing-and-troubleshooting.md)

[Tienda Office]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
