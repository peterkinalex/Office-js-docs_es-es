
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a>Transferir localmente complementos de Office en Office Online para pruebas

Puede instalar un complemento de Office para realizar pruebas sin tener que colocarlo primero en un catálogo de complementos con una instalación de prueba. La instalación de prueba se puede realizar en Office 365 o en Office Online. El procedimiento es ligeramente distinto para las dos plataformas. 

Cuando se transfiere localmente un complemento, el manifiesto de este se almacena en el almacenamiento local del navegador, por lo que si se quiere borrar el caché del navegador o cambiar a un navegador diferente, el complemento se debe transferir localmente de nuevo.


 >**Nota**  Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Transferir localmente complementos de Outlook para pruebas](sideload-outlook-add-ins-for-testing.md).


## <a name="sideload-an-office-add-in-on-office-365"></a>Sideload an Office Add-in on Office 365


1. Inicie sesión en su cuenta de Office 365.
    
2. Abra el iniciador de aplicaciones en el extremo izquierdo de la barra de herramientas y seleccione **Excel**, **Word** o **PowerPoint** y, después, cree un documento.
    
3. Abra la pestaña **Insertar** en la cinta de opciones y, en la sección **Complementos**, elija **Complementos de Office**.
    
4. En el cuadro de diálogo **Complementos de Office**, seleccione la pestaña **MI ORGANIZACIÓN** y, después, **Cargar mis complementos**.
    
    ![Cuadro de diálogo con el título Complemento de Office y un vínculo cerca del extremo superior izquierdo y el texto "Cargar mi complemento".](../../images/0e49f780-019a-4d97-9310-0eaddfa0c4dc.png)

5.  **Busque** en el archivo de manifiesto de complementos y, después, seleccione **Cargar**.
    
    ![Cuadro de diálogo de carga del complemento con los botones para examinar, cargar y cancelar.](../../images/039aef16-b12f-4d01-ad46-f13e01dd3162.png)

6. Verify that your complemento is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.
    

## <a name="sideload-an-office-add-in-on-office-online"></a>Transferir localmente un complemento de Office en Office Online




1. Open [Microsoft Office Online](https://office.live.com/).
    
2. En **Comenzar a trabajar con las aplicaciones en línea**, elija **Excel**, **Word** o **PowerPoint** y, después, abra un documento nuevo.
    
3. Abra la pestaña **Insertar** en la cinta de opciones y, en la sección **Complementos**, elija **Complementos de Office**.
    
4. En el cuadro de diálogo **Complementos de Office**, seleccione la pestaña **MIS COMPLEMENTOS**, elija **Administrar mis complementos** y, después, **Cargar mi complemento**.
    
    ![Cuadro de diálogo de Complementos de Office con una lista desplegable en la parte superior derecha con el texto "Administrar mis complementos" y otra más abajo con la opción "Cargar mi complemento".](../../images/d630d9d1-7dd5-45e3-860d-0ab069882548.png)

5.  **Busque** en el archivo de manifiesto de complementos y, después, seleccione **Cargar**.
    
    ![Cuadro de diálogo de carga del complemento con los botones para examinar, cargar y cancelar.](../../images/039aef16-b12f-4d01-ad46-f13e01dd3162.png)

6. Compruebe que el complemento está instalado. Por ejemplo, si es un comando de complemento, aparecerá en la cinta o en el menú contextual. Si es un complemento de panel de tareas, se mostrará el panel.
    
