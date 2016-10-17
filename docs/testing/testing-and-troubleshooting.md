
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Solucionar errores de usuario con los complementos de Office

A veces, los usuarios pueden experimentar problemas con los complementos de Office que desarrolla. Por ejemplo, un complemento no se puede cargar o está inaccesible. Use la información de este artículo como ayuda para resolver problemas comunes que experimentan los usuarios con el complemento de Office. 

También puede usar [Fiddler](http://www.telerik.com/fiddler) para identificar y depurar problemas con los complementos.

Después de resolver el problema del usuario, puede [responder directamente a las opiniones del cliente en la Tienda Office](https://msdn.microsoft.com/library/jj635874.aspx).

## <a name="common-errors-and-troubleshooting-steps"></a>Pasos para solucionar problemas y errores comunes

En la siguiente tabla se enumeran los mensajes de error comunes que los usuarios pueden encontrar, así como los pasos que los usuarios pueden realizar para resolver los errores.



|**Mensaje de error**|**Resolución**|
|:-----|:-----|
|Error de aplicación: el catálogo no está disponible|Compruebe la configuración del firewall."Catálogo" se refiere a la Tienda Office. Este mensaje indica que el usuario no puede tener acceso a la Tienda Office.|
|ERROR DE APLICACIÓN: No se pudo iniciar esta aplicación. Cierre este cuadro de diálogo para omitir el problema o haga clic en "Reiniciar" para volver a intentarlo.|Compruebe que están instaladas las últimas actualizaciones de Office o descargue la [actualización de Office 2013](https://support.microsoft.com/en-us/kb/2986156/).|
|Error: El objeto no admite la propiedad o el método 'defineProperty'|Confirme que Internet Explorer no se está ejecutando en modo de compatibilidad. Vaya a Herramientas >  **Configuración de Vista de compatibilidad**.|
|Lo sentimos, no pudimos cargar la aplicación porque la versión de su explorador no es compatible. Haga clic aquí para ver una lista de las versiones de explorador compatibles.|Asegúrese de que el explorador admite el almacenamiento local HTML5 o restablezca su configuración de Internet Explorer.Para obtener información sobre los exploradores compatibles, vea [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).|

## <a name="outlook-add-in-doesn't-work-correctly"></a>El complemento de Outlook no funciona correctamente

Si un complemento de Outlook que se ejecuta en Windows no funciona correctamente, pruebe a activar la depuración de scripts en Internet Explorer. 


- Vaya a Herramientas >  **Opciones de Internet** > **Opciones avanzadas**.
    
- En  **Examinar**, desactive  **Deshabilitar la depuración de scripts (Internet Explorer)** y **Deshabilitar la depuración de scripts (otros)**.
    
Se recomienda desactivar estas opciones solo para solucionar el problema. Si las deja desactivadas, recibirá mensajes mientras navegue. Una vez resuelto el problema, vuelva a activar  **Deshabilitar depuración de scripts (Internet Explorer)** y **Deshabilitar depuración de scripts (otros)**.


## <a name="add-in-doesn't-activate-in-office-2013"></a>El complemento no se activa en Office 2013

Si el complemento no se activa cuando el usuario realiza los pasos siguientes:


1. Inicia sesión con su cuenta de Microsoft en Office 2013.
    
2. Habilita la comprobación de dos pasos para su cuenta de Microsoft.
    
3. Comprueba su identidad cuando se le solicita al intentar insertar un complemento.
    
Compruebe que están instaladas las últimas actualizaciones de Office o descargue la [actualización de Office 2013](https://support.microsoft.com/en-us/kb/2986156/).


## <a name="additional-resources"></a>Recursos adicionales



- [Depurar complementos en Office Online](../testing/debug-add-ins-in-office-online.md)
    
- [Transferir localmente un complemento de Office a iPad y Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Depurar complementos de Office en dispositivos iPad y Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
- [Crear y depurar complementos de Office en Visual Studio](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
    
- [Implementar e instalar complementos de Outlook para probarlos](../outlook/testing-and-tips.md)
    
