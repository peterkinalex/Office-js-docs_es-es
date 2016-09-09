# Método UI.messageParent

Entrega un mensaje desde el cuadro de diálogo a su pagina primaria o de apertura. La página que llama a esta API debe estar en el mismo dominio que la página primaria. 

## Sintaxis

```js
Office.context.ui.messageParent("Message from Dialog box");
```

## Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|messageObject|Cadena o booleano|Acepta un mensaje del cuadro de diálogo para entregarlo al complemento.|

## Valores devueltos
void

## Ejemplos
Para obtener ejemplos, vea el tema [Método DisplayDialogAsync](officeui.displaydialogasync.md).

