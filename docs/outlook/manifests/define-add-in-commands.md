# <a name="define-add-in-commands-in-your-manifest"></a>Definir comandos de complemento en el manifiesto

Los comandos de complemento proporcionan una manera sencilla de personalizar la interfaz de usuario predeterminada de Office con elementos de la interfaz de usuario que realizan acciones (por ejemplo, puede agregar botones personalizados a la cinta). Para crear comandos, agregue un nodo **[VersionOverrides](../../../reference/manifest/versionoverrides.md)** a un manifiesto existente. 

Cuando un manifiesto contiene el elemento **VersionOverrides**, las versiones de Word, Excel, Outlook y PowerPoint que admiten comandos de complemento usarán la información dentro de dicho elemento para cargar el complemento. Las versiones anteriores de productos de Office que no admiten comandos de complemento omitirán el elemento.

Cuando las aplicaciones cliente reconocen el nodo **VersionOverrides**, el nombre del complemento aparece en la cinta de opciones y no en un panel tareas o de lectura y redacción. El complemento no aparece en ambos lugares.
 
## <a name="versionoverrides"></a>VersionOverrides

El elemento [VersionOverrides](../../../reference/manifest/versionoverrides.md) es el elemento raíz que contiene la información de los comandos de complemento implementados por el complemento. Se admite en el esquema de manifiesto v1.1 y posterior.

Existen dos versiones del esquema **VersionOverrides**.

| Versión del esquema | Descripción |
|----------------|-------------|
| 1.0 | Admite comandos de complemento para las versiones de escritorio de aplicaciones de Office. | 
| 1.1 | Agrega compatibilidad para los [paneles de tareas anclables](./pinnable-taskpane.md) y los complementos móviles. **Nota:** Actualmente, solo se admite en Outlook 2016 para Windows y Outlook para iOS |

Un complemento puede admitir varias versiones del esquema **VersionOverrides** anidando versiones más recientes dentro de la versión anterior. Esto permite que los clientes admitan versiones más recientes para aprovecharse de las características nuevas, mientras se permite que los clientes más antiguos carguen la versión anterior. Para obtener información, vea [Implementar varias versiones](../../../reference/manifest/versionoverrides.md#implementing-multiple-versions).

El elemento **VersionOverrides** incluye los siguientes elementos secundarios:

- [Descripción](../../../reference/manifest/description.md)
- [Requisitos](../../../reference/manifest/requirements.md)
- [Hosts](../../../reference/manifest/hosts.md)
- [Recursos](../../../reference/manifest/resources.md)
- [VersionOverrides](../../../reference/manifest/versionoverrides.md)

En el diagrama siguiente se muestra la jerarquía de los elementos que se han usado para definir comandos de complemento. 

![Jerarquía de los elementos de comandos de complementos en el manifiesto](../../../images/080da303-51c4-4882-b74a-7ba11517c0ad.png)

## <a name="sample-manifests"></a>Manifiestos de ejemplo

Para ver un manifiesto de ejemplo que implementa los comandos de complemento para Word, Excel y PowerPoint, consulte [Simple add-in commands sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple) (Ejemplo sencillo de comandos de complemento).

Para ver un manifiesto de ejemplo que implementa comandos de complemento para Outlook, consulte [Sample manifest file for an Outlook add-in](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml) (Ejemplo de archivo de manifiesto para un complemento de Outlook).

## <a name="additional-resources"></a>Recursos adicionales

- [Comandos de complemento para Outlook](../../outlook/add-in-commands-for-outlook.md)
    
- [Manifiestos de complementos de Outlook](../../outlook/manifests/manifests.md)
    
- [Ejemplo de demostración de comando de complemento de Outlook](https://github.com/jasonjoh/command-demo)
