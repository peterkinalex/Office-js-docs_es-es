

# userProfile

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

### Miembros

####  displayName :String

Obtiene el nombre para mostrar del usuario.

##### Tipo:

*   String

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### Ejemplo

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  emailAddress :String

Obtiene la dirección de correo electrónico SMTP del usuario.

##### Tipo:

*   String

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### Ejemplo

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  timeZone :String

Obtiene la zona horaria predeterminada del usuario.

##### Tipo:

*   String

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### Ejemplo

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```