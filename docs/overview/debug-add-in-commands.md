# Usar el registro de tiempo de ejecución para depurar comandos de complemento

Los clientes de escritorio de Office 16 tienen una nueva característica disponible para registrar la información útil. Entre otras cosas, esta herramienta puede ayudarle a diagnosticar los errores en el manifiesto del complemento que resulta especialmente útil si está creando manifiestos con comandos de complementos. 

La documentación completa de la característica está en camino pero, mientras tanto, así es cómo se puede usar para depurar problemas al analizar los manifiestos con comandos de complementos.

##Activar el registro de tiempo de ejecución

**Importante**: El registro de tiempo de ejecución tiene un **acierto de rendimiento**. Actívelo solamente cuando necesite depurar problemas en los complementos.

1. Asegúrese de que haya una versión que admita el registro en tiempo de ejecución. Necesita clientes de **escritorio de Office 16** con una versión igual o mayor que **16.0.7019**
2. Agregue la clave de registro `RuntimeLogging` dentro de `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`. 
3. Establezca el valor predeterminado de la clave como la ruta de acceso completa del archivo en el que quiere que se escriba el registro. Consulte [clave del registro de muestra](RuntimeLogging/EnableRuntimeLogging.zip) (descomprimir)

El registro debe tener este aspecto: ![](http://i.imgur.com/Sa9TyI6.png)

Si necesita desactivar la función, simplemente quite la clave del registro. 

##Diagnosticar problemas de comandos
El registro de tiempo de ejecución es útil para detectar **problemas con su manifiesto** que son difíciles de detectar, por ejemplo, incongruencia entre identificadores de recursos, longitudes no válidas, que la validación de esquema XSD no detecta. 

Estos son los pasos que puede probar:
 
1. Siga las instrucciones en el [archivo Léame](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/README.md) para cargar el complemento en paralelo. 
2. Si no ve el proyecto de botones de la cinta de opciones y no aparece nada en el cuadro de diálogo de complementos, compruebe los registros
3. Busque el identificador de su complemento, que define en su manifiesto, para buscar mensajes que pertenezcan a ese complemento. Los registros informan de este identificador como `SolutionId`. Se recomienda que solo transfiera localmente un complemento a la vez para evitar la aparición de gran cantidad de mensajes que no pertenecen al complemento. 

En el ejemplo siguiente, RuntimeLogging ayudó a identificar un control que apunta a un archivo de recurso que no existe. La solución es corregir el error ortográfico (si existe) o realmente agregar el recurso que falta.

![](http://i.imgur.com/f8bouLA.png) 

##Problemas conocidos relacionados con el registro
El registro de tiempo de ejecución aún tiene errores conocidos. Puede ver varios mensajes que están clasificados de manera inapropiada o confusa. Por ejemplo:

- Los mensajes `Medium  Current host not in add-in's host list` seguidos de `Unexpected Parsed manifest targeting different host` se clasifican de forma incorrecta. No son errores y se pueden omitir con seguridad.
- El mensaje `Unexpected   Add-in is missing required manifest fields  DisplayName` no contiene SolutionId del complemento infractor. Pero lo más probable es que esto NO esté relacionado con el complemento que se está depurando. 
- Todos los mensajes `Monitorable` son errores esperados desde un punto de vista del sistema. A veces pueden indicar un problema con el manifiesto (como un elemento mal escrito que se ha omitido pero no ha hecho que el manifiesto dejara de procesarse). 

