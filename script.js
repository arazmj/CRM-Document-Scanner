function acquire() {
    disableDivs('BabCorner');    
    var image = scanImage();

    
    if (image != null) {
        attachImageToLetter(image);
    }

    hideCDivs();

    var saveloc = window.location;    
    window.location = "";
    window.location = saveloc;
}


function getTempDir() {
    var WshShell = new ActiveXObject("WScript.Shell");
    keyValue = WshShell.regRead("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders\\Cache");
    return keyValue;
}

function showImages() {
	alert('show images inside ');
}

function S4() {
    return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
}

function guid() {
    return (S4() + S4() + "-" + S4() + "-" + S4() + "-" + S4() + "-" + S4() + S4() + S4());
}

function Base64Encode(sText)
{
    var oXML =  new ActiveXObject('Msxml2.DOMDocument.3.0');  
    var oNode = oXML.createElement('base64');
    oNode.dataType = 'bin.base64';
    oNode.nodeTypedValue = sText;  //Stream_StringToBinary(sText);
    return oNode.text;       
}

function ReadBinaryFile(FileName){
    var adTypeBinary = 1;
    var BinaryStream = new ActiveXObject("ADODB.Stream");
    BinaryStream.Type = adTypeBinary;

    BinaryStream.open();
    BinaryStream.loadFromFile(FileName);
    return BinaryStream.read();
}

function GenerateScannedImageFileName() {
    var currentTime = new Date();
    var month = currentTime.getMonth();
    var day = currentTime.getDate();
    var year = currentTime.getFullYear();    
    return 'ScannedImage ' + month + '-' + day + '-' + year + '.jpg';
}


function attachImageToLetter(imageBase64) {
    var xml = "" +
    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
    "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">" +
    "  <soap:Header>" +
    "    <CrmAuthenticationToken xmlns=\"http://schemas.microsoft.com/crm/2007/WebServices\">" +
    "      <AuthenticationType xmlns=\"http://schemas.microsoft.com/crm/2007/CoreTypes\">0</AuthenticationType>" +
    "      <OrganizationName xmlns=\"http://schemas.microsoft.com/crm/2007/CoreTypes\">" + ORG_UNIQUE_NAME + "</OrganizationName>" +
    "      <CallerId xmlns=\"http://schemas.microsoft.com/crm/2007/CoreTypes\">00000000-0000-0000-0000-000000000000</CallerId>" +
    "    </CrmAuthenticationToken>" +
    "  </soap:Header>" +
    "  <soap:Body>" +
    "    <Create xmlns=\"http://schemas.microsoft.com/crm/2007/WebServices\">" +
    "      <entity xsi:type=\"annotation\">" +
    "        <documentbody>" + imageBase64 + "</documentbody>" +
    "        <filename>" + GenerateScannedImageFileName() + "</filename>" +
    "        <mimetype>image/jpeg</mimetype>" +
    "        <notetext>This image has been scanned using Gostareh Negar CRM Image Acquisition</notetext>" +
    "        <objectid type=\"" + crmForm.ObjectTypeName + "\">" + crmForm.ObjectId.toString() + "</objectid>" +
    "        <objecttypecode>" + crmForm.ObjectTypeName + "</objecttypecode>" +
    "        <subject>Test Subject</subject>" +
    "      </entity>" +
    "    </Create>" +
    "  </soap:Body>" +
    "</soap:Envelope>" +
    "";    

    var xmlHttpRequest = new ActiveXObject("Msxml2.XMLHTTP");

    xmlHttpRequest.Open("POST", "/mscrmservices/2007/CrmService.asmx", false);
    xmlHttpRequest.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/crm/2007/WebServices/Create");
    xmlHttpRequest.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
    xmlHttpRequest.setRequestHeader("Content-Length", xml.length);
    xmlHttpRequest.send(xml);

    var resultXml = xmlHttpRequest.responseXML;
}


function Form_Onload(controlToReplace) {
    var script = document.createElement('script');
    script.type = 'text/javascript';
    script.src = '/ISV/Gostareh Negar/CRM Document Scanner/script.js';
    document.getElementsByTagName('head')[0].appendChild(script);
    

    if (crmForm.ObjectId == null) {
        document.getElementById(controlToReplace).style.visibility="hidden"
        return;
    }
    
    var xml = "" +
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">" +
            "  <soap:Header>" +
            "    <CrmAuthenticationToken xmlns=\"http://schemas.microsoft.com/crm/2007/WebServices\">" +
            "      <AuthenticationType xmlns=\"http://schemas.microsoft.com/crm/2007/CoreTypes\">0</AuthenticationType>" +
            "      <OrganizationName xmlns=\"http://schemas.microsoft.com/crm/2007/CoreTypes\">" + ORG_UNIQUE_NAME + "</OrganizationName>" +
            "      <CallerId xmlns=\"http://schemas.microsoft.com/crm/2007/CoreTypes\">00000000-0000-0000-0000-000000000000</CallerId>" +
            "    </CrmAuthenticationToken>" +
            "  </soap:Header>" +
            "  <soap:Body>" +
            "    <Fetch xmlns=\"http://schemas.microsoft.com/crm/2007/WebServices\">" +
            "      <fetchXml>&lt;fetch mapping='logical'&gt;&lt;entity name='annotation'&gt;&lt;attribute name='annotationid' /&gt;&lt;attribute name='filename' /&gt;&lt;attribute name='mimetype' /&gt;&lt;attribute name='objectid' /&gt;&lt;filter&gt;&lt;condition attribute='objectid' operator='eq' value='" + crmForm.ObjectId.toString() + "' /&gt;&lt;/filter&gt;&lt;/entity&gt;&lt;/fetch&gt;</fetchXml>" +
            "    </Fetch>" +
            "  </soap:Body>" +
            "</soap:Envelope>" +
            "";

    var xmlHttpRequest = new ActiveXObject("Msxml2.XMLHTTP");

    xmlHttpRequest.Open("POST", "/mscrmservices/2007/CrmService.asmx", false);
    xmlHttpRequest.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/crm/2007/WebServices/Fetch");
    xmlHttpRequest.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
    xmlHttpRequest.setRequestHeader("Content-Length", xml.length);
    xmlHttpRequest.send(xml);

    var resultXml = xmlHttpRequest.responseXML.text;
    xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
    xmlDoc.async = false;
    xmlDoc.loadXML(resultXml);


    ////////////////////////////////////////////////////////////////////////////
    

    // create the div tag        
    var div = document.createElement('div');
    div.setAttribute('style', 'border-style: double; width = 100');    
    // create images
    var images = new Array();
    var imagesNumber = Number(xmlDoc.getElementsByTagName('annotationid').length.valueOf());
    for (var i = 0; i < imagesNumber; i++) {
        //xmlDoc.getElementsByTagName('result')[0].getElementsByTagName('mimetype')[0] == null
        var annotationId = xmlDoc.getElementsByTagName('result')[i].getElementsByTagName('annotationid')[0].text;
        var mimetype = xmlDoc.getElementsByTagName('result')[i].getElementsByTagName('mimetype')[0];
             
        if (mimetype == null) {
            images[i] = null;
            continue;
        }
        var link = document.createElement('a');
        link.setAttribute('href', '/' + ORG_UNIQUE_NAME + '/Activities/Attachment/download.aspx?AttachmentType=5&AttachmentId=' + annotationId);
        link.setAttribute('type', 'image/png');
        var image = document.createElement('img');        
        image.src = '/' + ORG_UNIQUE_NAME + '/Activities/Attachment/download.aspx?AttachmentType=5&AttachmentId=' + annotationId;
        image.width = Number(240);
        image.height = Number(360);
        link.appendChild(image);
        images[i] = link;
    }

    // append images to the div
    for (var i = 0; i < imagesNumber; i++) {            
        if (images[i] != null) {
            div.appendChild(images[i]);
        }
    }

    // replace the control with the div
    document.getElementById(controlToReplace).parentNode.replaceChild(div, document.getElementById(controlToReplace)); 
}

function scanImage() {    
    try {        
        var CommonDialog = new ActiveXObject("WIA.CommonDialog");
        imageFile = CommonDialog.ShowAcquireImage(1, 0, 65536, "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}", false, true, false);

        if (imageFile != null) {            
            var file = getTempDir() + "\\" + guid();
            imageFile.SaveFile(file);
            var binImage = ReadBinaryFile(file)
            return Base64Encode(binImage);
        }
        else {
            return null;
        }
    } catch (e) {
        alert(e.message);
    }
}

cDivs = new Array();
function disableDivs() {
    d = document.getElementsByTagName("BODY")[0];
    for (x = 0; x < arguments.length; x++) {
        if (document.getElementById(arguments[x])) {
            xPos = document.getElementById(arguments[x]).offsetLeft;
            yPos = document.getElementById(arguments[x]).offsetTop;
            oWidth = document.getElementById(arguments[x]).offsetWidth;
            oHeight = document.getElementById(arguments[x]).offsetHeight;
            cDivs[cDivs.length] = document.createElement("DIV");
            //cDivs[cDivs.length]             
            cDivs[cDivs.length - 1].style.width = oWidth + "px";
            cDivs[cDivs.length - 1].style.height = oHeight + "px";
            cDivs[cDivs.length - 1].style.position = "absolute";
            cDivs[cDivs.length - 1].style.left = xPos + "px";
            cDivs[cDivs.length - 1].style.top = yPos + "px";
            cDivs[cDivs.length - 1].style.backgroundColor = "#999999";
            cDivs[cDivs.length - 1].style.opacity = .6;
            cDivs[cDivs.length - 1].style.filter = "alpha(opacity=60)";
            d.appendChild(cDivs[cDivs.length - 1]);
        }
    }
}

function hideCDivs() {
    for (hippopotamus = 0; hippopotamus < cDivs.length; hippopotamus++) {
        document.getElementsByTagName("BODY")[0].removeChild(cDivs[hippopotamus]);
    }
    cDivs = [];
}