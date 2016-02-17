# Crear el primer complemento de Excel

_Se aplica a: Excel 2016, Office 2016_

Los siguientes pasos le guiarán a lo largo del proceso de creación de un sencillo complemento del panel de tareas que carga datos en una hoja de cálculo y crea un gráfico básico en Excel 2016.

![Complemento de informe de ventas trimestrales](images/QuarterlySalesReport_report.PNG)


Comenzará creando una aplicación web con HTML y jQuery. A continuación, creará un archivo de manifiesto XML que especifica dónde desea ubicar su aplicación web y cómo debe aparecer en Excel. 


### Escribir código

1- Cree una carpeta en la unidad local denominada QuarterlySalesReport (por ejemplo, C:\\QuarterlySalesReport). Guarde en esta carpeta todos los archivos creados en los pasos siguientes.

2- Cree la página HTML que se cargará en el complemento del panel de tareas. Asigne al archivo el nombre **Home.HTML** y pegue el código siguiente en el archivo.

```html
	
	<!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Quarterly Sales Report</title>   

        <script src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>

        <link href="Office.css" rel="stylesheet" type="text/css" />

        <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js" type="text/javascript"></script>

        <link href="Common.css" rel="stylesheet" type="text/css" />
        <script src="Notification.js" type="text/javascript"></script>

        <script src="Home.js" type="text/javascript"></script>

        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">

    </head>
    <body class="ms-font-m">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>This sample shows how to load some sample data into the worksheet, and then create a chart using the Excel JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="load-data-and-create-chart">Click me!</button>
            </div>
        </div>
    </body>
    </html>	

```  

3- Cree un archivo denominado **Common.css** para almacenar los estilos personalizados y pegue el código siguiente en el archivo.

```css
	/* Common app styling */

    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; /* Fixed header height */
        overflow: hidden; /* Disable scrollbars for header */
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px; /* Same value as #content-header's height */
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; /* Enable scrollbars within main content section */
    }

    .padding {
        padding: 15px;
    }

    #notification-message {
        background-color: #818285;
        color: #fff;
        position: absolute;
        width: 100%;
        min-height: 80px;
        right: 0;
        z-index: 100;
        bottom: 0;
        display: none; /* Hidden until invoked */
    }

        #notification-message #notification-message-header {
            font-size: medium;
            margin-bottom: 10px;
        }

        #notification-message #notification-message-close {
            background-image: url("../Images/Close.png");
            background-repeat: no-repeat;
            width: 24px;
            height: 24px;
            position: absolute;
            right: 5px;
            top: 5px;
            cursor: pointer;
        }

	
```

4- Cree un archivo para que contenga la lógica de programación del complemento en jQuery. Asigne al archivo el nombre **Home.js** y pegue el siguiente script en el archivo.
	
```js
    
    (function () {
        "use strict";

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

                $('#load-data-and-create-chart').click(loadDataAndCreateChart);
            });
        };

        // Load some sample data into the worksheet and then create a chart
        function loadDataAndCreateChart() {
            // Run a batch operation against the Excel object model
            Excel.run(function (ctx) {

                // Create a proxy object for the active worksheet
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();

                //Queue commands to set the report title in the worksheet
                sheet.getRange("A1").values = "Quarterly Sales Report";
                sheet.getRange("A1").format.font.name = "Century";
                sheet.getRange("A1").format.font.size = 26;

                //Create an array containing sample data
                var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
                              ["Frames", 5000, 7000, 6544, 4377],
                              ["Saddles", 400, 323, 276, 651],
                              ["Brake levers", 12000, 8766, 8456, 9812],
                              ["Chains", 1550, 1088, 692, 853],
                              ["Mirrors", 225, 600, 923, 544],
                              ["Spokes", 6005, 7634, 4589, 8765]];

                //Queue a command to write the sample data to the specified range
                //in the worksheet and bold the header row
                var range = sheet.getRange("A2:E8");
                range.values = values;
                sheet.getRange("A2:E2").format.font.bold = true;

                //Queue a command to add a new chart
                var chart = sheet.charts.add("ColumnClustered", range, "auto");

                //Queue commands to set the properties and format the chart
                chart.setPosition("G1", "L10");
                chart.title.text = "Quarterly sales chart";
                chart.legend.position = "right"
                chart.legend.format.fill.setSolidColor("white");
                chart.dataLabels.format.font.size = 15;
                chart.dataLabels.format.font.color = "black";
                var points = chart.series.getItemAt(0).points;
                points.getItemAt(0).format.fill.setSolidColor("pink");
                points.getItemAt(1).format.fill.setSolidColor('indigo');

                //Run the queued commands, and return a promise to indicate task completion
                return ctx.sync();
            })
              .then(function () {
                  app.showNotification("Success");
                  console.log("Success!");
              })
            .catch(function (error) {
                // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
```


5- Cree un archivo para que contenga la lógica de programación para proporcionar notificaciones en el complemento en caso de error. Esto resulta útil durante la depuración. Asigne al archivo el nombre **Notification.js** y pegue el siguiente script en el archivo.
	
```js
    
    /* Notification functionality */

    var app = (function () {
        "use strict";

        var app = {};

        // Initialization function (to be called from each page that needs notification)
        app.initialize = function () {
            $('body').append(
                '<div id="notification-message">' +
                    '<div class="padding">' +
                        '<div id="notification-message-close"></div>' +
                        '<div id="notification-message-header"></div>' +
                        '<div id="notification-message-body"></div>' +
                    '</div>' +
                '</div>');

            $('#notification-message-close').click(function () {
                $('#notification-message').hide();
            });


            // After initialization, expose a common notification function
            app.showNotification = function (header, text) {
                $('#notification-message-header').text(header);
                $('#notification-message-body').text(text);
                $('#notification-message').slideDown('fast');
            };
        };

        return app;
    })();
```

6- Cree un archivo de manifiesto XML para especificar dónde se encuentra la aplicación web y cómo desea que aparezca en Excel. Asigne al archivo el nombre **QuarterlySalesReportManifest.xml** y pegue el siguiente código XML en el archivo.
    
	```xml
	<?xml version="1.0" encoding="UTF-8"?>
    <!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
      <Id>ab2991e7-fe64-465b-a2f1-c865247ef434</Id>
      <Version>1.0.0.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-US</DefaultLocale>
      <DisplayName DefaultValue="Quarterly Sales Report Sample" />
      <Description DefaultValue="Quarterly Sales Report Sample"/>
      <Capabilities>
        <Capability Name="Workbook" />
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="\\MyShare\QuarterlySalesReport\Home.html" />
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
	```

7- Genere un GUID usando un generador en línea de su elección. A continuación, reemplace con ese GUID el valor de la etiqueta **Id** que se muestra en el paso anterior. 

8-	Guarde todos los archivos. Acaba de escribir su primer complemento de Excel. 

### Pruébelo

La forma más sencilla de implementar y probar el complemento consiste en copiar los archivos en un recurso compartido de red.

1- Cree una carpeta en un recurso compartido de red (por ejemplo, \\\MyShare\\QuarterlySalesReport) y copie todos los archivos en esa carpeta.  

2- Edite el elemento **SourceLocation** del archivo de manifiesto para que apunte a la ubicación del recurso compartido para la página .html del paso 1. 

3- Copie el manifiesto (QuarterlySalesReportManifest.xml) en un recurso compartido de red (por ejemplo, \\\MyShare\\MyManifests).

4- Agregue la ubicación del recurso compartido que contiene el manifiesto como un catálogo de aplicaciones de confianza en Excel.

      a-  Launch Excel and open a blank spreadsheet.  
    
      b-  Choose the **File** tab, and then choose **Options**.
    
      c-  Choose **Trust Center**, and then choose the **Trust Center Settings** button.
    
      d-  Choose **Trusted Add-in Catalogs**.
    
      e-  In the **Catalog Url** box, enter the path to the network share you created in step 3, and then choose **Add Catalog**.
    
      f-  Select the **Show in Menu** check box, and then choose **OK**. A message appears to inform you that your settings will be applied the next time you start Office. 
        
5- Pruebe y ejecute el complemento. 

      a-  On the **Insert tab** in Excel 2016, choose **My Add-ins**. 
      
      b-  In the **Office Add-ins** dialog box, choose **Shared Folder**.
      
      c-  Choose **Quarterly Sales Report Sample**>**Insert**. The add-in opens in a task pane to the right of the current worksheet, as shown in the following figure.
 
 ![Complemento de informe de ventas trimestrales](images/QuarterlySalesReport_taskpane.PNG)
   
      d-  Click the **Click me!** button to render the data and the chart inside the worksheet, as shown in the following figure.  To see the chart update dynamically, just change the data in the range. 
        
![Complemento de informe de ventas trimestrales](images/QuarterlySalesReport_report.PNG)


### Recursos adicionales
 

*  [Introducción a la programación de complementos de Excel](excel-add-ins-programming-overview.md)
*  [Explorador de fragmentos de código para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Ejemplos de código de complementos de Excel](excel-add-ins-code-samples.md) 
*  [Referencia de la API de JavaScript de complementos de Excel](excel-add-ins-javascript-reference.md)
