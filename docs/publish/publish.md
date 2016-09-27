
# Implementar y publicar el complemento de Office


Puede usar uno de varios métodos a la hora de implementar el complemento de Office para pruebas o para distribuirlo a los usuarios:

- [Transferencia local](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md): úsela como parte del proceso de desarrollo para probar la ejecución del complemento en Windows, Office Online, iPad o Mac.
- [Catálogo de SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md): úselo como parte del proceso de desarrollo para probar el complemento o para distribuir el complemento a los usuarios de su organización.
- [Vista previa del Centro de administración de Office 365](https://support.office.com/en-ie/article/Deploy-Office-Add-Ins-in-Office-365-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE): úsela para distribuir el complemento a los usuarios de su organización.
- [Tienda Office]: úsela para distribuir el complemento a los usuarios de forma pública.

Las opciones disponibles dependen del tipo de complemento que cree y del host de Office al que esté destinado.

### Opciones de implementación de complementos para Word, Excel y PowerPoint

| Punto de extensión            | Transferencia local | Catálogo de SharePoint | Vista previa del Centro de administración de Office 365 | Tienda Office |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| Contenido         | X           | X                  | X                               | X            |
| Panel de tareas       | X           | X                  | X                               | X            |
| Comando         | X           |                    | X                               | X            |

> **NOTA:** No se admiten los catálogos de SharePoint en Office 2016 para Mac. Para implementar complementos de Office en clientes Mac, debe enviarlos a la [Tienda Office].    

### Opciones de implementación para complementos de Outlook

| Punto de extensión     | Transferencia local | Servidor Exchange | Tienda Office |
|:---------|:-----------:|:---------------:|:------------:|
| Aplicación Correo | X           | X               | X            |
| Comando  | X           | X               | X            |

Para ampliar el alcance de su complemento, asegúrese de que funciona en distintas plataformas. Los complementos de Office son compatibles con Windows, Mac, Web, iOS y Android. Para obtener una vista general de las características compatibles con cada plataforma, vea [Office Add-in host and platform availability].   

Para obtener información sobre las licencias de los complementos de la Tienda Office, consulte [Licencias de complementos](https://msdn.microsoft.com/EN-US/library/office/jj163257.aspx).

Para obtener información sobre cómo pueden los usuarios finales adquirir, insertar y ejecutar complementos, consulte [Empezar a usar el complemento de Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

## Recursos adicionales

- [Disponibilidad de plataformas y hosts de los complementos de Office]
- [Implementar e instalar complementos de Outlook para probarlos](../outlook/testing-and-tips.md) 
- [Enviar complementos y aplicaciones web a la Tienda Office][#tienda-de-office]
- [Instrucciones de diseño para complementos de Office](../design/add-in-design)
- [Crear complementos de la Tienda Office efectivos](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Solucionar errores de usuario con los complementos de Office](../testing/testing-and-troubleshooting.md)

[Tienda Office]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Disponibilidad de plataformas y hosts de los complementos de Office]: http://dev.office.com/add-in-availability
