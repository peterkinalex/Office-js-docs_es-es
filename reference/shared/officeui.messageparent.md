# <a name="ui.messageparent-method"></a>Método UI.messageParent

Entrega un mensaje desde el cuadro de diálogo a su pagina primaria o de apertura. La página que llama a esta API debe estar en el mismo dominio que la página primaria. 

## <a name="syntax"></a>Sintaxis

```js
Office.context.ui.messageParent("Message from Dialog box");
```

## <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|messageObject|Cadena o booleano|Acepta un mensaje del cuadro de diálogo para entregarlo al complemento.|

## <a name="returns"></a>Valores devueltos
void

## <a name="examples"></a>Ejemplos
Para obtener ejemplos, vea el tema [Método DisplayDialogAsync](officeui.displaydialogasync.md).

