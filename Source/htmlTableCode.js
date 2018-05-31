// Generates the comparision code
function CreateSyntax(checkpointName, objectName, reportDifference)
{
  var call;
  call = Syntax.CreateInvoke();
  call.ClassValue = "HTMLTableCheckpoint";
  call.InvokeName = "Compare";
  call.IsProperty = false;
  call.AddParameter(checkpointName); 
  call.AddParameter(eval(objectName));
  call.AddParameter(reportDifference); 

  var condition;
  condition = Syntax.CreateCondition();
  condition.OperatorType = condition.otNot;
  condition.Right = call;

  call = Syntax.CreateInvoke();
  call.ClassValue = "Log";
  call.InvokeName = "Error";
  call.AddParameter("The \"" + checkpointName + "\" HTML table checkpoint failed.");
  call.IsProperty = false;

  var ifSyntax;
  ifSyntax = Syntax.CreateIf();
  ifSyntax.Condition = condition;
  ifSyntax.TrueSyntax = call;
                                    
  return ifSyntax;
}


// Record-time action
function RecordExecute()
{
  var htcForm = UserForms.htcTableSelectorForm;
  htcForm.ResetForm();
  if (htcForm.ShowModal() != mrOk)
    return;
    
  var checkpointName = htcForm.cxTextEditCheckpointName.Text;
  var objectName = GetSelectedItem(htcForm.cxListBoxObjectName); 
  var reportDifference = htcForm.cxCheckBoxReportDifference.Checked;
  
  Recorder.AddSyntaxToScript(CreateSyntax(checkpointName, objectName, reportDifference));
}

// Desing-time action
function DesignTimeExecute()
{
  var htcForm = UserForms.htcTableSelectorForm;
  htcForm.ResetForm();
  if (htcForm.ShowModal() != mrOk)
    return;
    
  var checkpointName = htcForm.cxTextEditCheckpointName.Text;
  var objectName = GetSelectedItem(htcForm.cxListBoxObjectName);
  var reportDifference = htcForm.cxCheckBoxReportDifference.Checked; 
  
  var script = Syntax.GenerateSource(CreateSyntax(checkpointName, objectName, reportDifference));
  
  var ctcForm= UserForms.ctcForm;
  ctcForm.ResetForm()

  ctcForm.cmScriptText.Lines.Text = script;
  ctcForm.ShowModal();
}

// Generates a list of the selected tables
function htcTableSelectorForm_ObjectPickerMain_OnObjectPicked(Sender)
{
  var expectedTagName = "TABLE";
  var pickedObject = eval(Sender.PickedObjectName);
  var htcForm = UserForms.htcTableSelectorForm;
  var htcListNames = htcForm.cxListBoxObjectName;
  
  if (!aqObject.IsSupported(pickedObject, "NativeWebObject") && !aqObject.IsSupported(pickedObject, "NativeFirefoxObject")) {
    htcListNames.Items.Text = "";
    return;
  }
  
  var isFireFox = aqObject.IsSupported(pickedObject, "NativeFirefoxObject");
  
  var page = GetPageObject(pickedObject, isFireFox);
  if (!isFireFox)
    var tables = IE_GetObjects(page, pickedObject, expectedTagName);
  else
    var tables = FireFox_GetObjects(page, pickedObject, expectedTagName);
  
  var tablesNames = new Array();
  for(var i = 0; i < tables.length; i++)
    if ("" == tables[i].MappedName)
      tablesNames[i] = tables[i].FullName;
    else
      tablesNames[i] = tables[i].MappedName;
 
  htcListNames.ScrollWidth = 2500;
  htcListNames.Items.Text = tablesNames.join("\r\n");
  htcListNames.ItemIndex = 0;
  htcForm.SetFocus();
}

// Get the parent page of a web control
function GetPageObject(childObject, isFireFox)
{
  var page = childObject;
  if (!isFireFox) {
    while(!(aqObject.IsSupported(page, "Type") && ("HTML Document" == page.Type)))
      page = page.Parent;
  }
  else {
    while(!aqObject.IsSupported(page, "contentDocument"))
      page = page.Parent;
  }
  return page;
}

// Get an array of child objects by tag name in Internet Explorer
function IE_GetObjects(document, pickedObject, tagName)
{
  if (aqObject.IsSupported(pickedObject, "Type") && ("HTML Document" == pickedObject.Type)) {
    var tables = VBArray(pickedObject.FindAllChildren("tagName", tagName, 20000)).toArray();
  }
  else { 
    var tables = new Array();
    var obj = pickedObject.Parent;
    while(!(aqObject.IsSupported(obj, "ObjectType") && ("Page" == obj.ObjectType))) {
      if (aqObject.IsSupported(obj, "tagName") && tagName == obj.tagName) {
        tables[tables.length] = obj;
        //break;
      }
      obj = obj.Parent;
    }
    if (pickedObject.tagName == tagName) {
      var tempTables = VBArray(document.FindAllChildren("sourceIndex", pickedObject.sourceIndex, 20000)).toArray();
      for(var j = 0; j < tempTables.length; j++)
        tables[tables.length] = tempTables[j];
    }
    var childrenTables = pickedObject.getElementsByTagName(tagName); 
    for(var i = 0; i < childrenTables.length; i++) {
      var tempTables = VBArray(document.FindAllChildren("sourceIndex", childrenTables[i].sourceIndex, 20000)).toArray();
      for(var j = 0; j < tempTables.length; j++)
        tables[tables.length] = tempTables[j];
    } 
  }
  return tables;
}

// Get an array of child objects by tag name in FireFox
function FireFox_GetObjects(document, pickedObject, tagName)
{
  if (aqObject.IsSupported(pickedObject, "contentDocument")) {
    var tables = VBArray(pickedObject.FindAllChildren("tagName", tagName, 20000)).toArray();
  }
  else {
    var tables = new Array();
    var obj = pickedObject.Parent;
    while(!aqObject.IsSupported(pickedObject, "contentDocument")) {
      if (aqObject.IsSupported(obj, "tagName") && tagName == obj.tagName) {
        tables[tables.length] = obj;
        break;
      }
      obj = obj.Parent;
    }
    if (aqObject.IsSupported(pickedObject, "tagName") && (tagName == pickedObject.tagName)) {
      var tempTables = VBArray(document.FindAllChildren("outerHTML", pickedObject.outerHTML, 20000)).toArray();
      for(var j = 0; j < tempTables.length; j++)
        tables[tables.length] = tempTables[j];
    }
    var childrenTables = pickedObject.getElementsByTagName(tagName); 
    for(var i = 0; i < childrenTables.length; i++) {
      var tempTables = VBArray(document.FindAllChildren("outerHTML", childrenTables.item(0).outerHTML, 20000)).toArray();
      for(var j = 0; j < tempTables.length; j++)
        tables[tables.length] = tempTables[j];
    }
  }
  return tables;
}

// Highlights the selected object on the screen
function htcTableSelectorForm_cxListBoxObjectName_OnClick(Sender)
{
  var selectedItem = GetSelectedItem(Sender);
  var pickedObject = eval(selectedItem);
  var color = 255 | (0 << 8) | (0 << 16);
  Sys.HighlightObject(pickedObject, 3, color); 
  var htcForm = UserForms.htcTableSelectorForm;
  htcForm.SetFocus();
}

// Create an HTMLTable checkpoint with the specified parameters
function htcTableSelectorForm_cxButtonOK_OnClick(Sender)
{
  var htcForm = UserForms.htcTableSelectorForm;
  
  var checkpointName = htcForm.cxTextEditCheckpointName.Text;
  if (checkpointName == "") {
    ShowMessage("Please specify a name for the checkpoint.");
    return;
  }            
  if (!aqUtils.IsValidIdent(checkpointName)) {
    ShowMessage("Please specify a valid checkpoint name.");
    return;
  } 

  if (!Stores.Exists("XML") && !Stores.Create("XML")) {
    ShowMessage("Failed to create a new XML item in the Stores collection");
    return;
  }

  if (XML.Contains(checkpointName)) {
    if (mrYes != aqDlg.MessageDlg("This project already contains the \"" + checkpointName + "\" XML object.\r\nDo you want to use it?", 
                mtWarning, mbYesNoCancel)) {
      return;
    }                     
  }
  
  var objectName = GetSelectedItem(htcForm.cxListBoxObjectName);
  if (objectName == "") {
    ShowMessage("Please select an object to extract data from.");
    return;
  }
  try {
    var object = eval(objectName);
  }
  catch(exception) {
    object = Utils.CreateStubObject();
  }
  if (!object.Exists) {
    ShowMessage("The object does not exist:\r\n" + objectName + "\r\nPlease select an existing object to extract data from.");
    return;
  }
  
  if (!LoadHTMLObjectToXMLCheckpoint(checkpointName, object)) {
    ShowMessage("Failed to load data from HTML object to XML checkpoint.");
    return;
  }
  
  htcForm.ModalResult = mrOk;
}

// Gets the selected combobox item
function GetSelectedItem(listbox)
{
  var itemsText = listbox.Items.Text;
  if ("" != itemsText) {
    var items = itemsText.split("\r\n");
    for(var i = 0; i < items.length; i++) {
      if (listbox.Selected(i))
        return items[i];
    }
  }
  return "";
}

// Shows the HTMLTable Checkpoint help topic
function htcTableSelectorForm_cxButtonHelp_OnClick(Sender)
{
  Help.ShowContext("HTMLTableCheckpointExtension.chm", 6002);
}

// Cancels the checkpoint creation
function htcTableSelectorForm_cxButtonCancel_OnClick(Sender)
{
  UserForms.htcTableSelectorForm.ModalResult = mrCancel;
}

// Copies the comparision code to the Clipboard
function ctcForm_cbCopy_OnClick(Sender)
{
	Sys.Clipboard = UserForms.ctcForm.cmScriptText.Lines.Text;
}

// Copies the full name of the picked item to a new line of the comparision code in the Clipboard dialog
function ctcForm_ObjectPicker_OnObjectPicked(Sender)
{
  var lines;

  lines = UserForms.ctcForm.cmScriptText.Lines; 
  lines.Text += "\r\n";
  lines.Text += UserForms.ctcForm.ObjectPicker.PickedObjectName;
}

// Closes the Clipboard dialog
function ctcForm_cbCancel_OnClick(Sender)
{
  UserForms.ctcForm.Hide();
}

// Shows a help topic for the Clipboard dialog 
function ctcForm_cbHelp_OnClick(Sender)
{
  Help.ShowContext("HTMLTableCheckpointExtension.chm", 6001);
}
  
// Load HTML object data to XML checkpoint
function LoadHTMLObjectToXMLCheckpoint(xmlCheckpointName, object)
{
  var xmlData = GetHTMLObjectAsXmlObject(object);        

  if (null == eval("XML." + xmlCheckpointName)) {
    var xmlOptions = XML.CreateCheckpointOptions();
    xmlOptions.CompareInSubTreeMode = false;
    xmlOptions.IgnoreAttributes = false;
    xmlOptions.IgnoreNamespaceDeclarations = true;
    xmlOptions.IgnoreNodeOrder = true;
    xmlOptions.IgnorePrefixes = false;
    if (!XML.CreateXML(xmlCheckpointName, xmlData, xmlOptions))
      return false;
  }
  return true;
}

// Compares the specified XML checkpoint and the specified HTML object
function CompareHTMLObject(xmlCheckpointName, object, reportDifference)
{
  var result = false;

  var xmlCheckpoint = eval("XML." + xmlCheckpointName);
  if (null == xmlCheckpoint) {
    Log.Error("The \"" + xmlCheckpointName + "\" XML checkpoint is not found.");
  }
  else if (9 != aqObject.GetVarType(object)) {
    Log.Error("The Object parameter must be an object.");
  }
  else {
    result = xmlCheckpoint.Compare(GetHTMLObjectAsXmlObject(object), reportDifference);
  }

  return result;
}

// Create an XML object based on the specified HTML object
function GetHTMLObjectAsXmlObject(object)
{
  var xmlData, currentNode;
  
  var storedProperties = new Array();
  storedProperties[0] = "data";  
  storedProperties[1] = "innerText";
  storedProperties[2] = "textContent";
  
  xmlData = Sys.OleObject("Msxml2.DOMDocument");
  xmlData.async = false;
  
  xmlData.loadXML("<?xml version='1.0'?><Body/>");
  currentNode = xmlData.selectSingleNode("//Body");
  
  HTMLObjectToXml(currentNode, object, storedProperties);
  return xmlData;
}
   
// Loads HTML object data to an XML object
function HTMLObjectToXml(xmlData, object, storedProperties)
{      
  if (aqObject.IsSupported(object, "tagName")) {
    switch(object.tagName) {
      case "TABLE":
        var currentNode = xmlData.ownerDocument.createNode(1, GetHTMLObjectID(object), "");
        xmlData.appendChild(currentNode);
        for(var i = 0; i < object.rows.length; i++) 
          HTMLObjectToXml(currentNode, object.rows.item(i), storedProperties);
        break;
        
      case "TR":
        var currentNode = xmlData.ownerDocument.createNode(1, "Row", "");
        xmlData.appendChild(currentNode);
        for(var i = 0; i < object.cells.length; i++) { 
          currentColumnName = i;
          var currentCellNode = xmlData.ownerDocument.createNode(1, "Col_" + GetValidXMLName(currentColumnName.toString().replace(/ +$/g, "")), "");
          currentNode.appendChild(currentCellNode);
          AddXMLAttribute(currentCellNode, "Column_Name", currentColumnName);
          HTMLObjectToXml(currentCellNode, object.cells.item(i), storedProperties);
        }
        break;
        
      case "TD":
        var currentNode = xmlData.ownerDocument.createNode(1, "CellContent", "");
        xmlData.appendChild(currentNode);
        for(var i = 0; i < object.childNodes.length; i++)
          HTMLObjectToXml(currentNode, object.childNodes.item(i), storedProperties);
        break;
        
      default: 
        for(var i = 0; i < object.childNodes.length; i++) {
          var currentNode = xmlData.ownerDocument.createNode(1, object.tagName + i, "");
          xmlData.appendChild(currentNode);
          AddObjectPropertiesAsAttributes(currentNode, object, storedProperties);
          HTMLObjectToXml(currentNode, object.childNodes.item(i), storedProperties);
        }
        break;
    }
  }
  else {
    var currentNode = xmlData.ownerDocument.createNode(1, "Object", "");
    xmlData.appendChild(currentNode);
    AddObjectPropertiesAsAttributes(currentNode, object, storedProperties);
  }
}

// Gets the HTML object ID
function GetHTMLObjectID(object)
{
  var id = "";
  if (aqObject.IsSupported(object, "NativeWebObject"))
    id = object.NativeWebObject.id;
  else if (aqObject.IsSupported(object, "NativeFirefoxObject"))
    id = object.NativeFirefoxObject.id;
  else if (aqObject.IsSupported(object, "id"))
    id = object.id;
  if ("" == id)
    id  = "AnID";
  return id;  
}

// Adds HTML object properties to XML data
function AddObjectPropertiesAsAttributes(xmlData, object, properties) 
{
  for(var i = 0; i < properties.length; i++) {
    if (aqObject.IsSupported(object, properties[i]))                                
      AddXMLAttribute(xmlData, properties[i], aqObject.GetPropertyValue(object, properties[i]));
  }
}

// Creates a new attribute in the XML object data
function AddXMLAttribute(xmlData, name, value)
{
  var attribute = xmlData.ownerDocument.createAttribute(GetValidXMLName(name));
  try {
    attribute.value = value;
  }
  catch(exception) {
      attribute.value = "UNKNOWN_VALUE_TYPE";
  }
  xmlData.attributes.setNamedItem(attribute);
}

// Generates a valid based on the specified value 
function GetValidXMLName(name)
{
  return name.replace(/\W/g, "_");
}