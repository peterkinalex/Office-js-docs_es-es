# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Validar y solucionar problemas con el manifiesto

Use estos métodos para validar y solucionar problemas del manifiesto. 

- [Validar el manifiesto de los complementos de Office con el esquema XML](validate-the-office-add-ins-manifest-against-the-xml-schema)
- [Use el registro de tiempo de ejecución para depurar el manifiesto del complemento de Office](use-runtime-logging-to-debug-the-manifest-for-your-office-add-in)

## <a name="validate-your-manifest-against-the-xml-schema"></a>Validar su manifiesto con el esquema XML

Para ayudar a asegurarse de que el archivo de manifiesto que describe el complemento de Office es correcto y está completo, valídelo de acuerdo con los archivos [definición de esquema XML (XSD)](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas). Puede utilizar una herramienta de validación de esquemas XML o [Visual Studio](../get-started/create-and-debug-office-add-ins-in-visual-studio.md) para validar el manifiesto. 

Para usar Visual Studio, vaya a Compilar > Publicar y elija **Realizar comprobaciones de validación**.

Para utilizar una herramienta de línea de comandos de validación de esquemas XML para validar el manifiesto:

1.  Instale [tar](https://www.gnu.org/software/tar/) y [libxml](http://xmlsoft.org/FAQ.html) si no lo ha hecho todavía. 
2.  Ejecute el comando siguiente. Reemplace XSD_FILE con la ruta de acceso al archivo XSD de manifiesto y XML_FILE con la ruta de acceso al archivo XML de manifiesto.

    xmllint --noout --schema XSD_FILE XML_FILE

## <a name="use-runtime-logging-to-debug-your-add-in-manifest"></a>Usar el registro de tiempo de ejecución para depurar el manifiesto del complemento

Puede usar el registro de tiempo de ejecución para depurar el manifiesto de su complemento. Esta característica puede ayudarle a identificar y corregir los problemas del manifiesto que no se detectan mediante la validación de esquema XSD, como identificadores de recursos que no coinciden. El registro de tiempo de ejecución es especialmente útil para depurar complementos que implementan comandos de complemento.  

>**Nota:** La característica de registro de tiempo de ejecución está disponible actualmente en Office 2016 para equipos de escritorio.

### <a name="turn-on-runtime-logging"></a>Activar el registro de tiempo de ejecución

>**Importante**: El registro de tiempo de ejecución afecta al rendimiento. Actívelo solamente cuando necesite depurar problemas en el manifiesto del complemento.

1. Asegúrese de que está ejecutando la compilación **16.0.7019** o posterior de Office 2016 para equipos de escritorio. 
2. Agregue la clave del registro `RuntimeLogging` bajo 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\'. 
3. Establezca el valor predeterminado de la clave como la ruta de acceso completa del archivo en el que quiere que se escriba el registro. Para obtener un ejemplo, vea [EnableRuntimeLogging.zip](RuntimeLogging/EnableRuntimeLogging.zip). 

 > **Nota:** El directorio en el que se escribirá el archivo de registro ya debe existir y debe tener permisos de escritura al mismo. 
 
La imagen siguiente muestra el aspecto que debe tener el registro. ![Captura de pantalla del editor del registro con una clave de registro RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

Para desactivar la función, quite la clave `RuntimeLogging` del registro. 

### <a name="troubleshoot-issues-with-your-manifest"></a>Solucionar problemas con el manifiesto

Para usar el registro de tiempo de ejecución para solucionar problemas al cargar un complemento:
 
1. [Transfiera localmente el complemento](sideload-office-add-ins-for-testing.md) para hacer pruebas. 

    >Nota: Recomendamos que solo transfiera localmente el complemento que va a probar para minimizar el número de mensajes en el archivo de registro.
2. Si no ocurre nada y no ve el complemento (y no aparece en el cuadro de diálogo Complementos), abra el archivo de registro.
3. Busque en el archivo de registro el identificador del complemento, que habrá definido en el manifiesto. En el archivo de registro, este identificador lleva la etiqueta `SolutionId`. 

En el ejemplo siguiente, el archivo de registro identifica un control que apunta a un archivo de recursos que no existe. En este ejemplo, la solución sería corregir el error en el manifiesto o agregar el recurso que falta.

![Captura de pantalla de un archivo de registro con una entrada que especifica un identificador de recurso que no se encuentra](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>Problemas conocidos con el registro de tiempo de ejecución

Puede ocurrir que los mensajes del archivo de registro resulten confusos o estén clasificados de forma incorrecta. Por ejemplo:

- El mensaje `Medium   Current host not in add-in's host list` seguido de `Unexpected Parsed manifest targeting different host` está clasificado incorrectamente como un error.
- Si ve el mensaje `Unexpected    Add-in is missing required manifest fields  DisplayName` y no contiene un identificador SolutionId, lo más probable es que el error no esté relacionado con el complemento que está depurando. 
- Todos los mensajes `Monitorable` son errores esperados desde un punto de vista del sistema. A veces indican un problema con el manifiesto, como un elemento mal escrito que se ha omitido pero no ha hecho que el manifiesto dejara de procesarse. 

## <a name="additional-resources"></a>Recursos adicionales

- [Manifiesto XML de complementos para Office](../overview/add-in-manifests.md)
- [Transferir localmente complementos de Office para pruebas](sideload-office-add-ins-for-testing.md)
- [Depurar complementos de Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

