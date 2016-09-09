
# LabsJS lab components

Labs.js proporciona cuatro tipos de componente que puede usar para montar su laboratorio. Cada tipo de componente admite un tipo específico de interacción de laboratorio (por ejemplo, problemas de varias opciones, problemas de respuesta libre o actividades, como ver páginas web en el iframe HTML de la lección).

## Componentes

Office Mix es compatible con los cuatro siguientes tipos de componentes de laboratorio: 


-  **Activity component** ( **IActivityComponent**). Presents the user with an activity that must be completed; for example, read a piece of text, watch a video, or interact with a simulation. For more information, see [Labs.Components.ActivityComponentInstance](../../../reference/office-mix/labs.components.activitycomponentinstance.md).
    
-  **Choice component** ( **IChoiceComponent**). Presents the user with a list of choices from which the user must select. Supports single or multiple responses (or no answer at all). Use this component type for true/false, multiple choice, multiple response, or polls. For more information, see [Labs.Components.ChoiceComponentInstance](../../../reference/office-mix/labs.components.choicecomponentinstance.md).
    
-  **Input component** ( **IInputComponent**). Enables free form user input. Use this component type when you want to get responses to questions or math problems from the user, for example, or for other problem types that require text inputs from the user. For more information, see [Labs.Components.InputComponentInstance](../../../reference/office-mix/labs.components.inputcomponentinstance.md).
    
-  **Dynamic component** ( **IDynamicComponent**). Generates other component types at runtime. Use this component type when you have branching questions, for example, where follow-up component types vary depending on a previous user input. This type also enables creating quiz banks or generating problems at runtime. For more information, see [Labs.Components.DynamicComponentInstance](../../../reference/office-mix/labs.components.dynamiccomponentinstance.md).
    

## Recursos adicionales



- [Complementos de Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Configurar y editar laboratorios de LabsJS para Office Mix](../../powerpoint/office-mix/configuring-and-editing-labsjs-labs-for-office-mix.md)
    
- [Tutorial: Crear su primer laboratorio para Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
