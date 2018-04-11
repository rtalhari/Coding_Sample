// ------------------------------------------------------------\\
// Client: Company A
// Entity: Appointment
// Description: Update address onChange of Regarding, if regarding is Account or Contact
// Created On: Mar 23, 2016
// Created By: Ricardo Dias
var AppointmentVersion = "1.0.0.1";
// 1.0.0.1 - 2017-08-29: Ricardo Dias - Implemented Composite field validation.
// ------------------------------------------------------------\\

var Appointment = {
    Regarding: "regardingobjectid",
    Location: "location",
};

/********** Helper Functions **********/
var path = Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc";

var CrmAttr = function (attributeName) {
    return Xrm.Page.getAttribute(attributeName);
};

var CrmComponent = function (componentId) {
    return Sys.Application.findComponent(componentId);
};

var CrmControl = function (controlId) {
    return Xrm.Page.getControl(controlId);
};

var CheckValue = function (attribute) {
    var hasValue = false;
    var value = CrmAttr(attribute);
    if ((null != value) && (null != value.getValue())) {
        hasValue = true;
    }
    return hasValue;
};
/********** Helper Functions **********/

function fillPatientFromOpportunity(id, odataSetName, isDirect) {
    var context = Xrm.Page.context;
    var serverUrl = context.getClientUrl();
    var patient_id;
    var ODATA_ENDPOINT = "/XRMServices/2011/OrganizationData.svc";
    $.ajax({
        type: "GET",
        async: false,
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: serverUrl + ODATA_ENDPOINT + "/" + odataSetName + "(guid'" + id + "')",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
        },
        success: function (data, textStatus, XmlHttpRequest) {
            patient_id = data.d.va_Patient.Id;
            patient_name = data.d.va_Patient.Name;
            fillPatient(patient_id);
        },
        error: function (XmlHttpRequest, textStatus, errorThrown) {
            //alert('Error : ' + errorThrown);
            alert("ERROR 1026: Script Exception Error.  Please contact your system administrator for assistance.");
        }
    });
}

function GetAddress1Composite(field, id) {
    var query;
    if (field == "account") {
        query = "" + path + "/AccountSet(guid'" + id + "')?$select=Address1_Line1,Address1_Line2,Address1_Line3,Address1_City,Address1_StateOrProvince,Address1_PostalCode,Address1_Country";
    } else if (field == "contact") {
        query = "" + path + "/ContactSet(guid'" + id + "')?$select=Address1_Line1,Address1_Line2,Address1_Line3,Address1_City,Address1_StateOrProvince,Address1_PostalCode,Address1_Country";
    }
    $.ajax({
        url: query.replace('{', '').replace('}', ''),
        type: "GET",
        contentType: "application/json; charset=utf-8",
        dataType: "json",        
        async: false,
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
        },
        success: function (data) {
            if (data !== null && data.d !== null && data.d.results !== null) {
                var Address1_Line1 = data.d.Address1_Line1 ? data.d.Address1_Line1 + " " : "";
                var Address1_Line2 = data.d.Address1_Line2 ? data.d.Address1_Line2 + " " : "";
                var Address1_Line3 = data.d.Address1_Line3 ? data.d.Address1_Line3 + " " : "";
                var Address1_City = data.d.Address1_City ? data.d.Address1_City + " " : "";
                var Address1_StateOrProvince = data.d.Address1_StateOrProvince ? data.d.Address1_StateOrProvince + " " : "";
                var Address1_PostalCode = data.d.Address1_PostalCode ? data.d.Address1_PostalCode + " " : "";
                var Address1_Country = data.d.Address1_Country ? data.d.Address1_Country : "";

                var location = Address1_Line1 + Address1_Line2 + Address1_Line3 + Address1_City + Address1_StateOrProvince + Address1_PostalCode + Address1_Country;                                
              
                CrmAttr(Appointment.Location).setValue(location);
            }
            else {
                CrmAttr(Appointment.Location).setValue(null);
            }
        },
         error: function (XmlHttpRequest, textStatus, errorThrown) {
            //alert('Error : ' + errorThrown);
            alert("ERROR 1026: Script Exception Error.  Please contact your system administrator for assistance.");
        }
    });
};

function CheckRegardingType() {
    var regardingObject = Xrm.Page.getAttribute("regardingobjectid");
    if (regardingObject != null && regardingObject.getValue() != null && regardingObject.getValue() != "undefined") {
        var entityType = regardingObject.getValue()[0].entityType;
        var entityId = regardingObject.getValue()[0].id;
        if (entityType == "account" || entityType == "contact") {
            GetAddress1Composite(entityType, entityId);
        }
    }
};

var Regarding_OnChange = function () {
    CheckRegardingType();
};

var Form_OnLoad = function () {
    if (Xrm.Page.ui.getFormType() == 1) {
        CheckRegardingType();
    }
};
