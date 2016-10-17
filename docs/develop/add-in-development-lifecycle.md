
# <a name="office-add-ins-development-lifecycle"></a>Ciclo de vida de desarrollo de complementos de Office


El ciclo de vida de desarrollo típico de un complemento para Office incluye los siguientes pasos:


1.  **Decida la finalidad del complemento.**
    
    Haga las siguientes preguntas:
    
      - ¿Qué utilidad tiene el complemento? 
    
      - ¿Cómo ayudará a los clientes a ser más productivos?
    
      - ¿En qué situaciones se usarán las características del complemento?
    

    Decida las características y los escenarios más importantes y concéntrese en el diseño circundante. 
    
2.  **Identifique los datos y el origen de datos del complemento.**
    
    ¿Los datos están en un documento, libro, presentación, proyecto o una base de datos de Access basada en el explorador, o son sobre un elemento o elementos en un buzón de Exchange Server o Exchange Online? ¿Los datos provienen de un origen externo como un servicio web?
    
3.  **Identifique el tipo de complemento y de aplicaciones host de Office que mejor se ajusta a la finalidad del complemento.**
    
    Tenga en cuenta lo siguiente para identificar los escenarios:
    
    - ¿Los clientes usarán el complemento para enriquecer el contenido de una base de datos basada en explorador de Access o de documento? Si es así, es aconsejable considerar la creación de un complemento de contenido. 
    
    - ¿Los clientes usarán el complemento al ver o redactar un mensaje de correo electrónico o una cita? ¿Es importante poder exponer el complemento según el contexto actual? ¿Es prioritario habilitar el complemento no solo en el escritorio, sino también en tabletas y teléfonos?
    
        Si su respuesta es afirmativa a cualquiera de estas preguntas, considere la posibilidad de crear un complemento de Outlook. Después, identifique el contexto que desencadenará el complemento (por ejemplo, cuando el usuario esté en un formulario de redacción, tipos de mensaje específicos, la presencia de datos adjuntos, dirección, sugerencia de tarea o sugerencia de reunión, o determinados patrones de cadena en el contenido de un correo electrónico o cita). Consulte [Activation rules for Outlook add-ins](../outlook/manifests/activation-rules.md) (Reglas de activación para complementos de Outlook) para averiguar cómo puede activar contextualmente el complemento de Outlook.
    
    - ¿Los clientes usarán el complemento para mejorar la experiencia de visualización o creación de un documento? Si es así, es aconsejable considerar la creación de un complemento de panel de tareas. 

    La compatibilidad de algunas API de complemento puede diferir entre las aplicaciones de Office y la plataforma en la que se ejecutan (Windows, Mac, web, móviles). Para ver la cobertura de API actual por cliente y plataforma, consulte nuestra página de [Office Add-in host and platform availability](https://dev.office.com/add-in-availability) (Disponibilidad de host y plataforma del complemento de Office).  
    
4.  **Diseñar e implementar la experiencia de usuario y la interfaz de usuario del complemento.**
    
    Diseñe una experiencia de usuario rápida y fluida que sea coherente, fácil de aprender y con escenarios para los que solo sea necesario llevar a cabo unos pocos pasos. Dependiendo de la finalidad del complemento, puede hacer uso de API o servicios web de terceros.
    
    Puede elegir entre una variedad de herramientas de desarrollo web y usar HTML y JavaScript para implementar la interfaz de usuario.
    
5.  **Cree un archivo de manifiesto XML basado en el esquema del manifiesto de Complementos de Office.**
    
    Cree un manifiesto XML para identificar el complemento y sus requisitos, especifique las ubicaciones del HTML y cualquier archivo JavaScript y CSS que use el complemento, y, dependiendo del tipo de complemento, el tamaño predeterminado y los permisos.
    
    Para los complementos de Outlook, puede especificar el contexto, en función del mensaje o la cita actual, en el que el complemento será relevante y según el cual Outlook activará la aplicación en la interfaz de usuario. También puede decidir qué dispositivos serán compatibles con el complemento. En el manifiesto, especifique el contexto como reglas de activación y los dispositivos compatibles.
    
6.  **Instalar y probar el complemento.**
    
    Coloque los archivos HTML y cualquier archivo JavaScript y CSS en los servidores web que se especifican en el archivo de manifiesto del complemento. El proceso de instalación dependerá del tipo de aplicación.
    
    Para complementos de Outlook, se debe instalar en un buzón de Exchange y especificar la ubicación del archivo de manifiesto del complemento en el Centro de administración de Exchange (EAC). Para obtener más información, consulte [Implementar e instalar complementos de Outlook para probarlos](../outlook/testing-and-tips.md).
    
7.  **Publicar el complemento.**
    
    Puede enviar el complemento a la Tienda Office desde la que los clientes pueden instalar el complemento. Además, puede publicar complementos de panel de tareas y de contenido en un catálogo de complementos de carpeta privada en SharePoint o en una carpeta de red compartida, y puede implementar un complemento de Outlook directamente para su organización. Para obtener información más detallada, consulte [Publicar el complemento de Office](../publish/publish.md).
    
8.  **Actualizar el complemento**
    
    Si el complemento llama a un servicio web y usted realiza actualizaciones en el servicio web después de publicar el complemento, no es necesario volver a publicar el complemento. Pero si cambia los elementos o datos que ha enviado para el complemento, como el manifiesto del complemento, las capturas de pantalla, los iconos o los archivos HTML o JavaScript, debe volver a publicar el complemento. En particular, si ha publicado el complemento en la Tienda Office, debe volver a enviar el complemento para que la Tienda Office pueda implementar dichos cambios. Debe volver a enviar el complemento con un manifiesto de complemento actualizado que incluya un nuevo número de versión. También debe asegurarse de actualizar el número de versión del complemento en el formulario de envío para que coincida con el número de versión del nuevo manifiesto. Para complementos de Outlook, debe asegurarse de que el elemento [Id](../../reference/manifest/id.md) contenga un UUID diferente en el manifiesto del complemento.
    
