# Elemento OfficeTab
Define la ficha de la cinta en la que aparece el comando del complemento. Puede ser la pestaña predeterminada (**Inicio**, **Mensaje** o **Reunión**), o una pestaña personalizada definida por el complemento. Se requiere este elemento.

## Elementos secundarios
|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  Grupo      | Sí |  Define un grupo de comandos. Solo se puede agregar un grupo por cada complemento a la ficha predeterminada.  |


Los siguientes valores son `id` valores de ficha válidos por host: Los valores en **negrita** son compatibles en el escritorio y en línea (por ejemplo, Word 2016 para Windows y Word Online). 

### Outlook 
- **TabDefault**

### Word
- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### Excel
- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval 

### PowerPoint
- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### OneNote
- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## Group
Un grupo de puntos de extensión de UI en una ficha. Un grupo puede tener hasta seis controles. El atributo **id** es obligatorio y cada **id** debe ser único en el manifiesto. El **id** es una cadena de 125 caracteres como máximo. Consulte el [elemento Group](./group.md).

## Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
