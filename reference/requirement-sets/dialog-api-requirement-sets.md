
# <a name="dialog-api-requirement-sets"></a>Conjuntos de requisitos de la API de cuadros de diálogo

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Los complementos de Office usan los conjuntos de requisitos especificados en el manifiesto o usan una comprobación en tiempo de ejecución para determinar si un host de Office admite las API necesarias para el complemento. Para obtener más información, consulte [Specify Office hosts and API requirements (Especificar hosts de Office y requisitos de la API)](../docs/overview/specify-office-hosts-and-api-requirements.md).

Los complementos de Office se ejecutan en varias versiones de Office. En la siguiente tabla se enumeran los conjuntos de requisitos de la API de cuadros de diálogo, las aplicaciones de host de Office que admiten ese conjunto de requisitos y la compilación o números de versión de la aplicación de Office.

|  Conjunto de requisitos  |  Office 2013 para Windows | Office 2016 para Windows*   |  Office 2016 para iPad  |  Office 2016 para Mac  | Office Online  | 
|:-----|-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Compilación 15.0.4855.1000 o posterior | Versión 1602 (compilación 6741.0000) o posterior | 1.22 o posterior | 15.20 o posterior| Estamos trabajando en ello. |

>&#42; **Nota:** El número de compilación para Office 2016 que se ha instalado mediante MSI es 16.0.4266.1001. Para usar la API de cuadros de diálogo, ejecute la actualización de Office para obtener la última versión. 

Para obtener más información sobre las versiones y números de compilación, consulte:

- [Números de versión y compilación de las versiones del canal de actualización para los clientes de Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [¿Qué versión de Office estoy usando?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Dónde puede encontrar el número de versión y de compilación de una aplicación de cliente de Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office
Para obtener información sobre los conjuntos de requisitos comunes de la API, consulte [Office common API requirement sets (Conjuntos de requisitos comunes de la API de Office)](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>API de cuadros de diálogo 1.1 
La API de cuadros de diálogo 1.1 es la primera versión de la API. Para obtener más información sobre la API, consulte los temas de referencia de [API de cuadros de diálogo](../shared/officeui.md).

## <a name="additional-resources"></a>Recursos adicionales

- [Especificar los hosts de Office y los requisitos de la API](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Manifiesto XML de complementos para Office](https://dev.office.com/docs/add-ins/overview/add-in-manifests)
