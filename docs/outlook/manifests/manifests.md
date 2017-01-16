
# <a name="outlook-add-in-manifests"></a>Manifiestos de complementos de Outlook

Un complemento de Outlook está formado por dos componentes: el manifiesto del complemento XML y una página web, compatible con la biblioteca de JavaScript para complementos de Office (office.js). En el manifiesto se describe cómo se integra el complemento en diferentes clientes de Outlook. Actualmente hay tres versiones del esquema del manifiesto, incluida **VersionOverrides**. Le recomendamos que use la versión 1.1 del esquema del manifiesto y **VersionOverrides** 1.0 para crear el complemento. A continuación encontrará un ejemplo.

 >**Nota**  Todos los valores de la dirección URL en el siguiente ejemplo comienzan con "YOUR_WEB_SERVER". Este valor es un marcador de posición. En un manifiesto válido real, estos valores tendrían direcciones URL web https válidas.

```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-80.png" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read taskpane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppCompose/Home/Home.html"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
        <Host Name="Mailbox" />
    </Hosts>
    <Requirements>
        <Sets>
            <Set Name="MailBox" MinVersion="1.1" />
        </Sets>
    </Requirements>
    <!-- These elements support older clients that don't support add-in commands -->
    <FormSettings>
        <Form xsi:type="ItemRead">
            <DesktopSettings>
                <!-- NOTE: Just reusing the read taskpane page that is invoked by the button
                     on the ribbon in clients that support add-in commands. You can
                     use a completely different page if desired -->
                <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html" />
                <RequestedHeight>450</RequestedHeight>
            </DesktopSettings>
        </Form>
        <Form xsi:type="ItemEdit">
            <DesktopSettings>
                <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppCompose/Home/Home.html" />
            </DesktopSettings>
        </Form>
    </FormSettings>
    <Permissions>ReadWriteItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    </Rule>
    <DisableEntityHighlighting>false</DisableEntityHighlighting>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

        <Requirements>
            <bt:Sets DefaultMinVersion="1.3">
                <bt:Set Name="Mailbox" />
            </bt:Sets>
        </Requirements>
        <Hosts>
            <Host xsi:type="MailHost">

                <DesktopFormFactor>
                    <FunctionFile resid="functionFile" />

                    <!-- Message compose form -->
                    <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="msgComposeDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="msgComposeFunctionButton">
                                    <Label resid="funcComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcComposeSuperTipTitle" />
                                        <Description resid="funcComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>addDefaultMsgToBody</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="msgComposeMenuButton">
                                    <Label resid="menuComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuComposeSuperTipTitle" />
                                        <Description resid="menuComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="msgComposeMenuItem1">
                                            <Label resid="menuItem1ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ComposeLabel" />
                                                <Description resid="menuItem1ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg1ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgComposeMenuItem2">
                                            <Label resid="menuItem2ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ComposeLabel" />
                                                <Description resid="menuItem2ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg2ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgComposeMenuItem3">
                                            <Label resid="menuItem3ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ComposeLabel" />
                                                <Description resid="menuItem3ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg3ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                                    <Label resid="paneComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneComposeSuperTipTitle" />
                                        <Description resid="paneComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="composeTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Appointment compose form -->
                    <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="apptComposeDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="apptComposeFunctionButton">
                                    <Label resid="funcComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcComposeSuperTipTitle" />
                                        <Description resid="funcComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>addDefaultMsgToBody</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="apptComposeMenuButton">
                                    <Label resid="menuComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuComposeSuperTipTitle" />
                                        <Description resid="menuComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="apptComposeMenuItem1">
                                            <Label resid="menuItem1ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ComposeLabel" />
                                                <Description resid="menuItem1ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg1ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptComposeMenuItem2">
                                            <Label resid="menuItem2ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ComposeLabel" />
                                                <Description resid="menuItem2ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg2ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptComposeMenuItem3">
                                            <Label resid="menuItem3ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ComposeLabel" />
                                                <Description resid="menuItem3ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg3ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="apptComposeOpenPaneButton">
                                    <Label resid="paneComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneComposeSuperTipTitle" />
                                        <Description resid="paneComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="composeTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Message read form -->
                    <ExtensionPoint xsi:type="MessageReadCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="msgReadDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="msgReadFunctionButton">
                                    <Label resid="funcReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcReadSuperTipTitle" />
                                        <Description resid="funcReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>getSubject</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="msgReadMenuButton">
                                    <Label resid="menuReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuReadSuperTipTitle" />
                                        <Description resid="menuReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="msgReadMenuItem1">
                                            <Label resid="menuItem1ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ReadLabel" />
                                                <Description resid="menuItem1ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemClass</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgReadMenuItem2">
                                            <Label resid="menuItem2ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ReadLabel" />
                                                <Description resid="menuItem2ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getDateTimeCreated</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgReadMenuItem3">
                                            <Label resid="menuItem3ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ReadLabel" />
                                                <Description resid="menuItem3ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemID</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                                    <Label resid="paneReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneReadSuperTipTitle" />
                                        <Description resid="paneReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="readTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Appointment read form -->
                    <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="apptReadDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="apptReadFunctionButton">
                                    <Label resid="funcReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcReadSuperTipTitle" />
                                        <Description resid="funcReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>getSubject</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="apptReadMenuButton">
                                    <Label resid="menuReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuReadSuperTipTitle" />
                                        <Description resid="menuReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="apptReadMenuItem1">
                                            <Label resid="menuItem1ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ReadLabel" />
                                                <Description resid="menuItem1ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemClass</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptReadMenuItem2">
                                            <Label resid="menuItem2ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ReadLabel" />
                                                <Description resid="menuItem2ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getDateTimeCreated</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptReadMenuItem3">
                                            <Label resid="menuItem3ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ReadLabel" />
                                                <Description resid="menuItem3ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemID</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="apptReadOpenPaneButton">
                                    <Label resid="paneReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneReadSuperTipTitle" />
                                        <Description resid="paneReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="readTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>
                </DesktopFormFactor>
            </Host>
        </Hosts>

        <Resources>
            <bt:Images>
                <!-- Blue icon -->
                <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png" />
                <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png" />
                <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png" />
                <!-- Red icon -->
                <bt:Image id="red-icon-16" DefaultValue="YOUR_WEB_SERVER/images/red-16.png" />
                <bt:Image id="red-icon-32" DefaultValue="YOUR_WEB_SERVER/images/red-32.png" />
                <bt:Image id="red-icon-80" DefaultValue="YOUR_WEB_SERVER/images/red-80.png" />
                <!-- Green icon -->
                <bt:Image id="green-icon-16" DefaultValue="YOUR_WEB_SERVER/images/green-16.png" />
                <bt:Image id="green-icon-32" DefaultValue="YOUR_WEB_SERVER/images/green-32.png" />
                <bt:Image id="green-icon-80" DefaultValue="YOUR_WEB_SERVER/images/green-80.png" />
            </bt:Images>
            <bt:Urls>
                <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html" />
                <bt:Url id="readTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html" />
                <bt:Url id="composeTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppCompose/TaskPane/TaskPane.html" />
            </bt:Urls>
            <bt:ShortStrings>
                <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
                <!-- Compose mode -->
                <bt:String id="funcComposeButtonLabel" DefaultValue="Insert default message" />
                <bt:String id="menuComposeButtonLabel" DefaultValue="Insert message" />
                <bt:String id="paneComposeButtonLabel" DefaultValue="Insert custom message" />

                <bt:String id="funcComposeSuperTipTitle" DefaultValue="Inserts the default message" />
                <bt:String id="menuComposeSuperTipTitle" DefaultValue="Choose a message to insert" />
                <bt:String id="paneComposeSuperTipTitle" DefaultValue="Enter your own text to insert" />

                <bt:String id="menuItem1ComposeLabel" DefaultValue="Hello World!" />
                <bt:String id="menuItem2ComposeLabel" DefaultValue="Add-in commands are cool!" />
                <bt:String id="menuItem3ComposeLabel" DefaultValue="Visit dev.outlook.com" />

                <!-- Read mode -->
                <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
                <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
                <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

                <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
                <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
                <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

                <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
                <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
                <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
            </bt:ShortStrings>
            <bt:LongStrings>
                <!-- Compose mode -->
                <bt:String id="funcComposeSuperTipDescription" DefaultValue="Inserts text into body of the message or appointment. This is an example of a function button." />
                <bt:String id="menuComposeSuperTipDescription" DefaultValue="Inserts your choice of text into body of the message or appointment. This is an example of a drop-down menu button." />
                <bt:String id="paneComposeSuperTipDescription" DefaultValue="Opens a pane where you can enter text to insert in the body of the message or appointment. This is an example of a button that opens a task pane." />

                <bt:String id="menuItem1ComposeTip" DefaultValue="Inserts Hello World! into the body of the message or appointment." />
                <bt:String id="menuItem2ComposeTip" DefaultValue="Inserts Add-in commands are cool! into the body of the message or appointment." />
                <bt:String id="menuItem3ComposeTip" DefaultValue="Inserts Visit dev.outlook.com into the body of the message or appointment." />

                <!-- Read mode -->
                <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
                <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
                <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

                <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
                <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
                <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
            </bt:LongStrings>
        </Resources>
    </VersionOverrides>
</OfficeApp>

```


## <a name="schema-versions"></a>Versiones de esquema

No todos los clientes de Outlook admiten las características más recientes a la vez. Además, algunos usuarios de Outlook tendrán una versión más antigua. Las versiones de esquema permiten a los desarrolladores compilar complementos que son compatibles con versiones anteriores, que usan las características más recientes cuando están disponibles, pero que siguen funcionando en versiones anteriores.

El elemento  **VersionOverrides** del manifiesto es un ejemplo de esto. Todos los elementos definidos dentro de **VersionOverrides** invalidarán al mismo elemento en la otra parte del manifiesto. Esto significa que, siempre que sea posible, Outlook usará lo que haya en la sección **VersionOverrides** para configurar el complemento. Sin embargo, si la versión de Outlook no admite una versión determinada de **VersionOverrides**, Outlook la omitirá y dependerá de la información del resto del manifiesto. 

Este enfoque significa que los desarrolladores no tienen que crear varios manifiestos, sino que deben definir todo en un archivo.

Las versiones actuales del esquema son:


|||
|:-----|:-----|
|Versión|Descripción|
|v1.0|Admite la versión 1.0 de la API de JavaScript para Office. En el caso de los complementos de Outlook, se admite el formulario de lectura. |
|v1.1|Admite la versión 1.1 de la API de JavaScript para Office y  **VersionOverrides**. En el caso de los complementos de Outlook, esto agrega compatibilidad con el formulario de redacción.|
|**VersionOverrides** 1.0|Admite versiones posteriores de la API de JavaScript para Office y es compatible con los comandos de complementos.|
En este artículo se tratarán los requisitos para un manifiesto 1.1. Incluso si el manifiesto de su complemento usa el elemento  **VersionOverrides**, sigue siendo importante incluir los elementos del manifiesto 1.1 para permitir que la aplicación funcione con clientes anteriores que no admiten  **VersionOverrides**.


## <a name="root-element"></a>Elemento raíz

El elemento raíz del manifiesto del complemento de Outlook es  **OfficeApp**. Este elemento también declara el espacio de nombres predeterminado, la versión del esquema y el tipo de complemento. Coloque los demás elementos en el manifiesto dentro de las etiquetas de apertura y cierre. El siguiente es un ejemplo del elemento raíz:


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest>

</OfficeApp>
```


## <a name="version"></a>Versión

Esta es la versión del complemento específico. Si un desarrollador actualiza algo en el manifiesto, también se debe incrementar la versión. De este modo, cuando se instale el nuevo manifiesto, se sobrescribirá el anterior y el usuario podrá disfrutar de las funciones nuevas. Si este complemento se envió al almacén, el nuevo manifiesto se tendrá que volver a enviar y validar. Así, los usuarios de este complemento obtendrán el nuevo manifiesto actualizado automáticamente en unas horas tras su aprobación.

Si los permisos solicitados del complemento cambian, se solicitará a los usuarios que actualicen y vuelvan a otorgar su consentimiento para el complemento. Si el administrador ha instalado este complemento para toda la organización, el administrador tendrá que volver a otorgar su consentimiento primero. Los usuarios seguirán viendo la funcionalidad anterior mientras tanto.


## <a name="versionoverrides"></a>VersionOverrides

El elemento **VersionOverrides** es la ubicación de la información para comandos de complementos. Para obtener más información sobre este elemento, consulte [Definir comandos de complementos en el manifiesto del complemento de Outlook](../../outlook/manifests/define-add-in-commands.md).


## <a name="localization"></a>Localización

Algunos aspectos del complemento deben estar localizados para distintas configuraciones regionales, como el nombre, la descripción y la dirección URL que se carga. Estos elementos se pueden localizar fácilmente especificando el valor predeterminado y, después, invalidaciones de configuración regional en el elemento **Recursos** dentro del elemento **VersionOverrides**. A continuación, se muestra cómo reemplazar una imagen, una dirección URL y una cadena:


```XML
<Resources>
    <bt:Images>
      <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
        <!-- add information for other locales -->

    <bt:Urls>
      <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
        <!-- add information for other locales -->

    <bt:ShortStrings> 
      </bt:String>
      <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
        <bt:Override Locale="ar-sa" Value="<add localized value here>" />
        <!-- add information for other locales -->
    </bt:ShortStrings>

  </Resources>
```

La referencia de esquema contiene toda la información sobre los elementos que se pueden localizar.


## <a name="hosts"></a>Hosts

Los complementos de Outlook especifican el elemento  **Hosts** como el siguiente.


```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

Esto es independiente del elemento  **Hosts** dentro del elemento **VersionOverrides**, que se describe en [Definir comandos de complementos en el manifiesto de complemento de Outlook](../../outlook/manifests/define-add-in-commands.md).


## <a name="requirements"></a>Requisitos

El elemento  **Requirements** especifica el conjunto de las API disponibles para el complemento. Para un complemento de Outlook, el conjunto de requisitos debe ser Mailbox y un valor de 1.1 o superior. Consulte la referencia de API para ver la última versión del conjunto de requisitos. Consulte [API de complementos de Outlook](../../outlook/apis.md) para obtener más información sobre los conjuntos de requisitos.

El elemento  **Requirements** también puede aparecer en el elemento **VersionOverrides**, lo que permite que el complemento especifique un requisito distinto cuando se carga en clientes que admiten  **VersionOverrides**.

El siguiente ejemplo usa el atributo  **DefaultMinVersion** del elemento **Sets** para necesitar office.js versión 1.1 o superior, y el atributo **MinVersion** del elemento **Set** para necesitar el conjunto de requisitos de buzón versión 1.1.




```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```


## <a name="form-settings"></a>Configuración de formulario

El elemento  **FormSettings** lo usan los clientes de Outlook anteriores, que solo admiten el esquema 1.1 y no admiten **VersionOverrides**. Con este elemento, los desarrolladores definen cómo se mostrará el complemento en estos clientes. Hay dos partes: **ItemRead** e **ItemEdit**.  **ItemRead** se usa para especificar cómo se muestra el complemento en los mensajes cuando el usuario lee mensajes y citas. **ItemEdit** describe cómo se muestra el complemento cuando el usuario está redactando una respuesta, un nuevo mensaje, una nueva cita o editando una cita, casos en los que es el organizador.

Estas configuraciones están relacionadas directamente con las reglas de activación en el elemento  **Rule**. Por ejemplo, si un complemento especifica que debe aparecer en un mensaje en modo redacción, debe especificarse un formulario  **ItemEdit**.

Para obtener más información, vea la [Referencia de esquemas para manifiestos de complementos de Office (versión 1.1)](../../overview/add-in-manifests.md).

## <a name="app-domains"></a>Dominios de aplicación

El dominio de la página de inicio del complemento que se especifica en el elemento  **SourceLocation** es el dominio predeterminado del complemento. Sin usar los elementos **AppDomains** y **AppDomain**, si el complemento intenta navegar a otro dominio, el explorador abrirá una nueva ventana fuera del panel de complementos. Para que el complemento pueda navegar a otro dominio dentro del panel de complementos, agregue un elemento  **AppDomains** e incluya cada dominio adicional en su propio subelemento **AppDomain** en el manifiesto del complemento.

El siguiente ejemplo especifica un dominio  `https://www.contoso2.com` como segundo dominio al que el complemento puede navegar en el panel de complementos:




```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

Los dominios de aplicación también son necesarios para habilitar el uso compartido de cookies entre la ventana emergente y el complemento que se ejecuta en el cliente enriquecido.


## <a name="permissions"></a>Permisos

El elemento  **Permissions** contiene los permisos necesarios para el complemento. En general, debe especificar el permiso mínimo que necesita el complemento según los métodos exactos que planea usar. Por ejemplo, un complemento de correo que se activa en los formularios de redacción y solo lee pero no escribe en propiedades de elemento como [item.requiredAttendees](../../../reference/outlook/Office.context.mailbox.item.md) y no llama a [mailbox.makeEwsRequestAsync](../../../reference/outlook/Office.context.mailbox.md) para obtener acceso a las operaciones de los servicios Web Exchange debe especificar el permiso **ReadItem**. Para ver información detallada sobre los permisos disponibles, consulte [Especificar permisos para el acceso de los complementos de Outlook al buzón del usuario](../../outlook/understanding-outlook-add-in-permissions.md).


**Modelo de permisos de cuatro niveles para complementos de correo**

![Modelo de permisos de cuatro niveles para esquema v1.1 de aplicaciones de correo](../../../images/olowa15wecon_Permissions_4Tier.png)
```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>

```


## <a name="activation-rules"></a>Reglas de activación

Las reglas de activación se especifican en el elemento  **Rule**. El elemento  **Rule** puede aparecer como un elemento secundario del elemento **OfficeApp** en manifiestos 1.1, y además como un elemento secundario **ExtensionPoint** en **VersionOverrides**. Consulte [Definir comandos de complementos en el manifiesto de complemento de Outlook](../../outlook/manifests/define-add-in-commands.md) para obtener información detallada sobre cómo usar este elemento en **VersionOverrides**.

Las reglas de activación se pueden usar para activar un complemento en función de una o varias de las siguientes condiciones en el elemento actualmente seleccionado.


- El tipo de elemento o clase de mensaje
    
- La presencia de un tipo específico de entidad conocida, como un número de teléfono o dirección
    
- Una coincidencia de expresión regular en el cuerpo, asunto o dirección de correo electrónico del remitente
    
- La presencia de datos adjuntos
    
Para obtener información detallada y ejemplos de reglas de activación, vea [Reglas de activación para complementos de Outlook](../../outlook/manifests/activation-rules.md).


## <a name="next-steps-add-in-commands"></a>Pasos siguientes: comandos de complementos


Después de definir un manifiesto básico, [defina comandos de complementos para el complemento](../../outlook/manifests/define-add-in-commands.md). Los comandos de complementos presentan un botón en la cinta para que los usuarios puedan activar el complemento de una forma sencilla e intuitiva. Para obtener más información, consulte [Comandos de complementos de Outlook](../../outlook/add-in-commands-for-outlook.md).

Para obtener un complemento de ejemplo que define los comandos de complemento, consulte [command-demo](https://github.com/jasonjoh/command-demo).

## <a name="additional-resources"></a>Recursos adicionales



- [Complementos de Outlook](../../outlook/outlook-add-ins.md)
    
- [Reglas de activación para complementos de Outlook](../../outlook/manifests/activation-rules.md)
    
- [Localización de complementos para Office](../../develop/localization.md)
    
- [Crear un complemento de correo para Outlook que se ejecute en equipos de escritorio, tabletas y dispositivos móviles (esquema v1.1)](http://msdn.microsoft.com/library/8d425fb3-8a7c-429d-87b3-8046e964b153%28Office.15%29.aspx)
    
- [Privacidad, permisos y seguridad para los complementos de Outlook](../../outlook/privacy-and-security.md)
    
- [API de complementos de Outlook](../../outlook/apis.md)
    
- [Manifiesto XML de complementos para Office](../../overview/add-in-manifests.md)
    
- [Referencia de esquema para manifiestos de Complementos de Office (v1.1)](../../overview/add-in-manifests.md)
    
- [Tipos de elementos y clases de mensajes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)
    
- [Instrucciones de diseño para complementos de Office](../../design/add-in-design.md)
    
- [Comprender los permisos de los complementos de Outlook](../../outlook/understanding-outlook-add-in-permissions.md)
    
- [Usar las reglas de activación de las expresiones regulares para mostrar un complemento de Outlook](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Coincidencia de cadenas en un elemento de Outlook como entidades conocidas](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
