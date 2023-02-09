/**
 * @NApiVersion 2.x
 * @NScriptType ClientScript
 * @NModuleScope SameAccount
 */
define(['N/url', 'N/https', 'N/log', 'N/record', 'N/ui/dialog', 'N/ui/message', 'N/format'],
/**
 * @param{https} https
 * @param{log} log
 * @param{record} record
 * @param{redirect} redirect
 * @param{dialog} dialog
 * @param{message} message
 */
function(url, https, log, record, dialog, message, format) {
    
    /**
     * Function to be executed after page is initialized.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.mode - The mode in which the record is being accessed (create, copy, or edit)
     *
     * @since 2015.2
     */
    function pageInit(scriptContext) {

    }

    /**
     * Function to be executed when field is changed.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.sublistId - Sublist name
     * @param {string} scriptContext.fieldId - Field name
     * @param {number} scriptContext.lineNum - Line number. Will be undefined if not a sublist or matrix field
     * @param {number} scriptContext.columnNum - Line number. Will be undefined if not a matrix field
     *
     * @since 2015.2
     */
    function fieldChanged(scriptContext) {
        var CurrentRecord = scriptContext.currentRecord;
        var dateFrom = CurrentRecord.getValue({
            fieldId: 'custpage_duedatefromflt'
        })||'';
        if(dateFrom != ''){
            dateFrom = format.format({
                value: dateFrom, 
                type: format.Type.DATE
            });
        }
        var dateTo = CurrentRecord.getValue({
            fieldId: 'custpage_duedatetoflt'
        })||'';
        if(dateTo != ''){
            dateTo = format.format({
                value: dateTo, 
                type: format.Type.DATE
            });
        }

        var account = CurrentRecord.getValue({
            fieldId: 'custpage_accountflt'
        })||'';
        var paymentType = CurrentRecord.getValue({
            fieldId: 'custpage_pymttypeflt'
        })||'';
        var paymentTypeText = CurrentRecord.getText({
            fieldId: 'custpage_pymttypeflt'
        })||'';
        var paymentBankName = CurrentRecord.getValue({
            fieldId: 'custpage_pymtbnkname'
        })||'';
        var paymentBankNameText = CurrentRecord.getText({
            fieldId: 'custpage_pymtbnkname'
        })||'';
        CurrentRecord.setValue({
            fieldId: 'custpage_pymttypetxt',
            value: paymentTypeText
        });
        var docType = CurrentRecord.getValue({
            fieldId: 'custpage_doctypeflt'
        })||'';
        var docTypeText = CurrentRecord.getText({
            fieldId: 'custpage_doctypeflt'
        })||'';
        CurrentRecord.setValue({
            fieldId: 'custpage_doctypetxt',
            value: docTypeText
        });
        var currency = CurrentRecord.getValue({
            fieldId: 'custpage_currencyflt'
        })||'';
        var pageIndex = CurrentRecord.getValue({
            fieldId: 'custpage_paginationfld'
        })||'';

        var root = location.protocol + '//' + location.host;
	    var suiteleturl = url.resolveScript({
            scriptId: 'customscript_jcs_aging_payable_report_sl',
            deploymentId: 'customdeploy_jcs_aging_payable_report_sl',
            returnExternalUrl: false 
        });  
	    var requrl = root + 
        suiteleturl + 
        '&custpage_accountflt='+account+
        '&custpage_duedatefromflt='+dateFrom+
        '&custpage_duedatetoflt='+ dateTo + 
        '&custpage_pymttypeflt='+paymentType+
        '&custpage_pymtbnkname='+paymentBankName+
        '&custpage_doctypeflt='+ docType + 
        '&custpage_pymttypetxt='+paymentTypeText+
        '&custpage_pymtbnknametxt='+paymentBankNameText+
        '&custpage_doctypetxt='+ docTypeText + 
        '&custpage_currencyflt='+currency+
        '&custpage_paginationfld='+ pageIndex;
        
	    var win = window.open(requrl, "_self");
    }

    /**
     * Function to be executed when field is slaved.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.sublistId - Sublist name
     * @param {string} scriptContext.fieldId - Field name
     *
     * @since 2015.2
     */
    function postSourcing(scriptContext) {

    }

    /**
     * Function to be executed after sublist is inserted, removed, or edited.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.sublistId - Sublist name
     *
     * @since 2015.2
     */
    function sublistChanged(scriptContext) {

    }

    /**
     * Function to be executed after line is selected.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.sublistId - Sublist name
     *
     * @since 2015.2
     */
    function lineInit(scriptContext) {

    }

    /**
     * Validation function to be executed when field is changed.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.sublistId - Sublist name
     * @param {string} scriptContext.fieldId - Field name
     * @param {number} scriptContext.lineNum - Line number. Will be undefined if not a sublist or matrix field
     * @param {number} scriptContext.columnNum - Line number. Will be undefined if not a matrix field
     *
     * @returns {boolean} Return true if field is valid
     *
     * @since 2015.2
     */
    function validateField(scriptContext) {

    }

    /**
     * Validation function to be executed when sublist line is committed.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.sublistId - Sublist name
     *
     * @returns {boolean} Return true if sublist line is valid
     *
     * @since 2015.2
     */
    function validateLine(scriptContext) {

    }

    /**
     * Validation function to be executed when sublist line is inserted.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.sublistId - Sublist name
     *
     * @returns {boolean} Return true if sublist line is valid
     *
     * @since 2015.2
     */
    function validateInsert(scriptContext) {

    }

    /**
     * Validation function to be executed when record is deleted.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.sublistId - Sublist name
     *
     * @returns {boolean} Return true if sublist line is valid
     *
     * @since 2015.2
     */
    function validateDelete(scriptContext) {

    }

    /**
     * Validation function to be executed when record is saved.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @returns {boolean} Return true if record is valid
     *
     * @since 2015.2
     */
    function saveRecord(scriptContext) {
        var CurrentRecord = scriptContext.currentRecord;
        var paymentTypeText = CurrentRecord.getText({
            fieldId: 'custpage_pymttypeflt'
        })||'';
        CurrentRecord.setValue({
            fieldId: 'custpage_pymttypetxt',
            value: paymentTypeText
        });
        var docTypeText = CurrentRecord.getText({
            fieldId: 'custpage_doctypeflt'
        })||'';
        CurrentRecord.setValue({
            fieldId: 'custpage_doctypetxt',
            value: docTypeText
        });
        return true;
    }

    return {
        /*pageInit: pageInit,*/
        fieldChanged: fieldChanged,
        /*postSourcing: postSourcing,
        sublistChanged: sublistChanged,
        lineInit: lineInit,
        validateField: validateField,
        validateLine: validateLine,
        validateInsert: validateInsert,
        validateDelete: validateDelete,*/
        saveRecord: saveRecord
    };
    
});
