
# Comandos de complementos para Excel, Word y PowerPoint

Los comandos de complementos son elementos de la interfaz de usuario que amplían la interfaz de usuario de Office e inician acciones en el complemento. Puede agregar un botón a la cinta, o bien un elemento a un menú contextual. Cuando el usuario selecciona un comando de complemento, se inician acciones como ejecutar código de JavaScript o mostrar una página del complemento en el panel de tareas. Los comandos de complementos ayudan a los usuarios a encontrar y usar su complemento, lo que puede aumentar la adopción y reutilización del complemento y mejorar la retención de clientes.

Para obtener información general sobre la característica, vea el vídeo [Add-in Commands in Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551) (Comandos de complementos en la cinta de Office).


**Complemento con comandos que se ejecuta en la versión para equipos de escritorio de Excel**
![Comandos de complementos](../../images/addincommands1.png)

**Complemento con comandos que se ejecuta en Excel Online**
![Comandos de complementos](../../images/addincommands2.png)

## Capacidades de comando
Actualmente, se admiten las capacidades de comando siguientes.

**Puntos de extensión**

- Pestañas de la cinta: ampliar las pestañas integradas o crear una pestaña personalizada.
- Menús contextuales: ampliar los menús contextuales seleccionados. 

**Tipos de controles**

- Botones sencillos: activan acciones específicas.
- Menús: contienen varios botones que activan acciones.

**Acciones**

- ShowTaskpane: muestra uno o varios paneles donde se cargan páginas HTML personalizadas.
- ExecuteFunction: carga una página HTML invisible y, después, ejecuta una función de JavaScript en la página. Para mostrar la interfaz de usuario dentro de su función (por ejemplo, errores, progreso, entrada adicional) puede usar la API [displayDialog](http://dev.office.com/reference/add-ins/shared/officeui).  

## Plataformas compatibles
Los comandos de complementos se admiten actualmente en las plataformas siguientes:

- Office 2016 para escritorio de Windows (versión 16.0.6769.0000 o posteriores)
- Office Online con cuentas personales
- Office Online con cuentas profesionales o educativas (vista previa)

Estarán disponibles en más plataformas próximamente.

## Introducción a los comandos de complementos

Para obtener más información sobre cómo especificar comandos de complementos en el manifiesto, consulte [Define add-in commands in your manifest](http://dev.office.com/docs/add-ins/outlook/manifests/define-add-in-commands) (Definir comandos de complementos en el manifiesto).

Para comenzar a usar comandos de complementos, vea los [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) (Ejemplos de comandos de complementos de Office) en GitHub.





