 

# <a name="understanding-api-requirement-sets"></a>Entender los conjuntos de requisitos de la API

Los complementos de Outlook declaran qué versiones de la API necesitan mediante el uso del elemento [Requirements](https://msdn.microsoft.com/EN-US/library/office/dn592036.aspx) de su [manifiesto](https://msdn.microsoft.com/en-us/library/office/fp123693.aspx). Los complementos de Outlook siempre incluyen un elemento [Set](https://msdn.microsoft.com/EN-US/library/office/dn592049.aspx) con un atributo `Name` establecido en `Mailbox` y un atributo `MinVersion` establecido en el conjunto de requisitos mínimos de la API que admite los escenarios de los complementos.

Por ejemplo, el siguiente fragmento de manifiesto indica un conjunto de requisitos mínimos de la versión 1.1:

```
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Todas las API de Outlook pertenecen al `Mailbox`[conjunto de requisitos](https://msdn.microsoft.com/EN-US/library/office/dn535871.aspx#SpecifyRequirementSets_intro). El conjunto de requisitos `Mailbox` tiene versiones y cada nuevo conjunto de API que publicamos pertenece a una versión posterior del conjunto. No todos los clientes de Outlook admiten el conjunto más reciente de API, pero si un cliente de Outlook declara compatibilidad para un conjunto de requisitos, será compatible con todas las API de ese conjunto de requisitos.

Al establecer una versión de conjunto de requisitos mínimos en el manifiesto se controla en qué cliente de Outlook aparecerá el complemento. Si un cliente no admite el conjunto de requisitos mínimos, no carga el complemento. Por ejemplo, si se especifica la versión 1.3 del conjunto de requisitos, esto significa que el complemento no se mostrará en ningún cliente de Outlook que no admita al menos la versión 1.3.

## <a name="using-apis-from-later-requirement-sets"></a>Usar las API desde conjuntos de requisitos posteriores

Al establecer un conjunto de requisitos no se limitan las API disponibles que puede usar el complemento. Por ejemplo, si el complemento especifica un conjunto de requisitos de la versión 1.1, pero está ejecutando un cliente de Outlook que admite 1.3, el complemento puede usar las API del conjunto de requisitos de la versión 1.3\.

Para usar las API más recientes, los desarrolladores pueden comprobar su existencia mediante la técnica estándar de JavaScript.

```
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Dichos controles no son necesarios para ninguna API que esté presente en la versión del conjunto de requisitos especificada en el manifiesto.

## <a name="choosing-a-minimum-requirement-set"></a>Elegir un conjunto de requisitos mínimos

Los desarrolladores deben usar el conjunto de requisitos más antiguo que contenga el conjunto fundamental de las API para su escenario, sin el que no funcionará el complemento.

## <a name="clients"></a>Clientes

Los siguientes clientes admiten complementos de Outlook.

| Client | Conjuntos admitidos de requisitos de la API |
| --- | --- |
| Outlook 2016 | 1.1, 1.2, 1.3 |
| Mac Outlook 2016 | 1.1 |
| Outlook 2013 | 1.1, 1.2, 1.3 |
| Outlook en la web (Office 365 y Outlook.com) | 1.1, 1.2, 1.3 |
| Outlook Web App (Exchange 2013 local) | 1.1 |
| Outlook Web App (Exchange 2016 local) | 1.1, 1.2. 1.3 |
>**Nota** La compatibilidad con la versión 1.3 de Outlook 2013 se ha agregado como parte de la [actualización del 8 de diciembre de 2015 para Outlook 2013 (KB3114349)](https://support.microsoft.com/en-us/kb/3114349)
