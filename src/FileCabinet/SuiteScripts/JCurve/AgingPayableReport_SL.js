/**
* @NApiVersion 2.1
* @NScriptType Suitelet
*/
/*
* @name:                                       AgingPayableReport_SL.js
* @author:                                     Kamlesh Patel
* @summary:                                    Script Description
* @copyright:                                  © Copyright by Jcurve Solutions
* Date Created:                                Tue Nov 03 2022 1:07:32 PM
* Change Logs:
* Date                          Author               Description
* Tue Nov 03 2022 1:07:32 PM -- Kamlesh Patel -- Initial Creation
*/
//Global Variable
let DefaultListSize = 25;

define(['N/ui/serverWidget', 'N/encode', 'N/error', 'N/file', 'N/format', 'N/log', 'N/record', 'N/runtime', 'N/search', './Library_CM.js', 'N/config'],
/**
* @param{encode} encode
* @param{error} error
* @param{file} file
* @param{format} format
* @param{log} log
* @param{record} record
* @param{runtime} runtime
* @param{search} search
*/
(ui, encode, error, file, format, log, record, runtime, search, lib, config) => {
    /**
    * Defines the Suitelet script trigger point.
    * @param {Object} scriptContext
    * @param {ServerRequest} scriptContext.request - Incoming request
    * @param {ServerResponse} scriptContext.response - Suitelet response
    * @since 2015.2
    */
    const onRequest = (scriptContext) => {
        var companyInfo = config.load({
            type: config.Type.COMPANY_INFORMATION
        });
        var companyTimezone = companyInfo.getValue({
            fieldId:'timezone'
        });
        let request = scriptContext.request;
        let response = scriptContext.response;
        let SSID = 'customsearch_jcs_aging_report_2';
        let searchObj = search.load({
            id: SSID
        });
        //Summary search for total of column in excel
        let SummarySSID = 'customsearch_jcs_aging_report_2_2';
        let summarySearchObj = search.load({
            id: SummarySSID
        });

        //Get VBN Notes List
        var VBNNotesDataList = {};
        var vbnNotesSearch = lib.runSearch(
            search.create({
                type: "customrecord_vbn_vendor_billing_notes",
                filters:
                [
                ],
                columns:
                [
                    search.createColumn({
                        name: "name",
                        sort: search.Sort.ASC,
                        label: "Name"
                    }),
                    search.createColumn({name: "id", label: "ID"}),
                    search.createColumn({name: "scriptid", label: "Script ID"}),
                    search.createColumn({name: "custrecord_vbn_nss_document_number", label: "NSS Document Number"}),
                    search.createColumn({name: "csegsub_cost_center", label: "Sub Cost Center"}),
                    search.createColumn({name: "custrecord_vbn_ar_account", label: "A/R Account"}),
                    search.createColumn({name: "custrecord_vbn_bill_status", label: "Bill Status"}),
                    search.createColumn({name: "custrecord_vbn_billing_note_date", label: "Billing Note Date"}),
                    search.createColumn({name: "custrecord_vbn_memo", label: "Memo"}),
                    search.createColumn({name: "custrecord_vbn_expected_pay_date", label: "EXPECTED PAY DATE"})
                ]
            })
            );

            for(var index = 0; index < vbnNotesSearch.length; index++){
                var CurrentRow = vbnNotesSearch[index];
                var VBN_No  = CurrentRow.getValue({name: "name", sort: search.Sort.ASC, label: "Name"})||'';
                var VBNBillingNotesDate  = CurrentRow.getValue({name: "custrecord_vbn_billing_note_date", label: "Billing Note Date"})||'';
                VBNNotesDataList[VBN_No] = {vbnbillingdate:VBNBillingNotesDate}
            }
            log.debug('VBNNotesDataList', JSON.stringify(VBNNotesDataList));

            if (scriptContext.request.method == 'GET'){
                try{
                    //create form
                    var form = ui.createForm({
                        title : 'Aging Payable Report'
                    });
                    //Field Group
                    var filterFieldsGroup = form.addFieldGroup({ id : 'fldgrp_filter',label : 'Filters' });
                    var resultGroup = form.addFieldGroup({ id : 'fldgrp_result',label : 'List' });

                    //Filterfields
                    var account = form.addField({
                        id : 'custpage_accountflt',
                        type : ui.FieldType.SELECT,
                        label : 'Account',
                        container : 'fldgrp_filter'
                    });

                    var duedate_from = form.addField({
                        id : 'custpage_duedatefromflt',
                        type : ui.FieldType.DATE,
                        label : 'Due Date From',
                        container : 'fldgrp_filter'
                    });
                    var duedate_to = form.addField({
                        id : 'custpage_duedatetoflt',
                        type : ui.FieldType.DATE,
                        label : 'Due Date To',
                        container : 'fldgrp_filter'
                    });

                    var paymentType = form.addField({
                        id : 'custpage_pymttypeflt',
                        type : ui.FieldType.MULTISELECT,
                        source : 'paymentmethod',
                        label : 'Payment Type',
                        container : 'fldgrp_filter'
                    });

                    var paymentBankName = form.addField({
                        id : 'custpage_pymtbnkname',
                        type : ui.FieldType.MULTISELECT,
                        source : 'customrecord_billing_note_bank_master',
                        label : 'Payment Bank Name',
                        container : 'fldgrp_filter'
                    });

                    var documentType = form.addField({
                        id : 'custpage_doctypeflt',
                        type : ui.FieldType.MULTISELECT,
                        source : 'customlist_ap_source_form',
                        label : 'Document Type',
                        container : 'fldgrp_filter'
                    });

                    var paymentTypeTxt = form.addField({
                        id : 'custpage_pymttypetxt',
                        type : ui.FieldType.TEXT,
                        label : 'Payment Type TXT',
                        container : 'fldgrp_filter'
                    });
                    var documentTypeTxt = form.addField({
                        id : 'custpage_doctypetxt',
                        type : ui.FieldType.TEXT,
                        label : 'Document Type TXT',
                        container : 'fldgrp_filter'
                    });
                    paymentTypeTxt.updateDisplayType({
                        displayType : ui.FieldDisplayType.HIDDEN
                    });
                    documentTypeTxt.updateDisplayType({
                        displayType : ui.FieldDisplayType.HIDDEN
                    });

                    var currency = form.addField({
                        id : 'custpage_currencyflt',
                        type : ui.FieldType.MULTISELECT,
                        source : 'currency',
                        label : 'Currency',
                        container : 'fldgrp_filter'
                    });

                    //Account list load
                    var accountSearchObj = search.create({
                        type: "account",
                        filters:
                        [
                            ["type","anyof","AcctPay"],
                            "AND",
                            ["isinactive","is","F"]
                        ],
                        columns:
                        [
                            search.createColumn({name: "internalid", label: "Internal ID"}),
                            search.createColumn({
                                name: "displayname",
                                sort: search.Sort.ASC,
                                label: "Display Name"
                            })
                        ]
                    });
                    var searchResultCount = accountSearchObj.runPaged().count;
                    log.debug("accountSearchObj result count",searchResultCount);
                    account.addSelectOption({
                        value : '',
                        text : '- All -'
                    });
                    accountSearchObj.run().each(function(result){
                        let value = result.getValue({name: "internalid", label: "Internal ID"});
                        let label = result.getValue({name: "displayname", sort: search.Sort.ASC, label: "Display Name"});
                        account.addSelectOption({
                            value : value,
                            text : label
                        });
                        return true;
                    });

                    //Pagination dropdown
                    var paginationFld = form.addField({
                        id : 'custpage_paginationfld',
                        type : ui.FieldType.SELECT,
                        label : 'List Range',
                        container : 'fldgrp_result'
                    });
                    var TotalResultCountFld = form.addField({
                        id : 'custpage_totalcountfld',
                        type : ui.FieldType.TEXT,
                        label : 'Total Results',
                        container : 'fldgrp_result'
                    });
                    TotalResultCountFld.updateDisplayType({
                        displayType : ui.FieldDisplayType.INLINE });

                        //add Sublist
                        var custList = form.addSublist({
                            id : 'custpage_resultsublist',
                            type : ui.SublistType.STATICLIST,
                            label : 'Result',
                            container : 'fldgrp_result'
                        });

                        //Set value that is already set in filter
                        var selectedAccount = request.parameters.custpage_accountflt||'';
                        form.updateDefaultValues({
                            custpage_accountflt : selectedAccount
                        });

                        var selectedFromDate = request.parameters.custpage_duedatefromflt||'';
                        form.updateDefaultValues({
                            custpage_duedatefromflt : selectedFromDate
                        });

                        var selectedToDate = request.parameters.custpage_duedatetoflt||'';
                        form.updateDefaultValues({
                            custpage_duedatetoflt : selectedToDate
                        });

                        var selectedPaymentType = request.parameters.custpage_pymttypeflt||'';
                        form.updateDefaultValues({
                            custpage_pymttypeflt : selectedPaymentType!=''?selectedPaymentType.split(','):selectedPaymentType
                        });

                       var selectedPaymentBankName = request.parameters.custpage_pymtbnkname||'';
                        form.updateDefaultValues({
                            custpage_pymtbnkname : selectedPaymentBankName!=''?selectedPaymentBankName.split(','):selectedPaymentBankName
                        });

                        var selectedDocType = request.parameters.custpage_doctypeflt||'';
                        form.updateDefaultValues({
                            custpage_doctypeflt : selectedDocType!=''?selectedDocType.split(','):selectedDocType
                        });

                        var selectedCurrency = request.parameters.custpage_currencyflt||'';
                        form.updateDefaultValues({
                            custpage_currencyflt :  selectedCurrency!=''?selectedCurrency.split(','):selectedCurrency
                        });

                        //add sublist fields

                        //add search Filters
                        /*if(selectedFromDate != ''){
                            searchObj.filters.push(search.createFilter({name: 'duedate', operator: search.Operator.ONORAFTER, values: selectedFromDate}));
                        }
                        if(selectedToDate != ''){
                            searchObj.filters.push(search.createFilter({name: 'duedate', operator: search.Operator.ONORBEFORE, values: selectedToDate}));
                        }*/
                        /*
                        OLD Formula before enahancement on 6 FEB 2023
                        selectedFromDate :: "CASE {type} WHEN 'Bill' THEN {duedate}WHEN 'Bill Credit' THEN {duedate}WHEN 'Journal' THEN {trandate} ELSE {trandate}  END"
                        selectedToDate :: "CASE {type} WHEN 'Bill' THEN {duedate}WHEN 'Bill Credit' THEN {duedate}WHEN 'Journal' THEN {trandate} ELSE {trandate}  END"
                        */

                        if(selectedFromDate != ''){
                            searchObj.filters.push(search.createFilter({
                                name: 'formuladate',
                                formula : "CASE {type} WHEN 'Bill' THEN {custbody_gisc_change_due}WHEN 'Bill Credit' THEN {custbody_gisc_change_due}WHEN 'Journal' THEN {trandate} ELSE {trandate}  END",
                                operator: search.Operator.ONORAFTER,
                                values: selectedFromDate
                            }));
                        }
                        if(selectedToDate != ''){
                            searchObj.filters.push(search.createFilter({
                                name: 'formuladate',
                                formula : "CASE {type} WHEN 'Bill' THEN {custbody_gisc_change_due}WHEN 'Bill Credit' THEN {custbody_gisc_change_due}WHEN 'Journal' THEN {trandate} ELSE {trandate}  END",
                                operator: search.Operator.ONORBEFORE,
                                values: selectedToDate
                            }));
                        }

                        if(selectedAccount != ''){
                            searchObj.filters.push(search.createFilter({name: 'account', operator: search.Operator.ANYOF, values: selectedAccount}));
                        }

                        if(selectedPaymentType != ''){
                            searchObj.filters.push(search.createFilter({name: 'custbody_vbn_paymentmethod', operator: search.Operator.ANYOF, values: selectedPaymentType.split(',')}));
                        }

                        if(selectedPaymentBankName != ''){
                            searchObj.filters.push(search.createFilter({name: 'custbody_bank_name', operator: search.Operator.ANYOF, values: selectedPaymentBankName.split(',')}));
                        }

                        if(selectedDocType != ''){
                            searchObj.filters.push(search.createFilter({name: 'custbody_sourcefrom', operator: search.Operator.ANYOF, values: selectedDocType.split(',')}));
                        }
                        if(selectedCurrency != ''){
                            searchObj.filters.push(search.createFilter({name: 'currency', operator: search.Operator.ANYOF, values: selectedCurrency.split(',')}))
                        }

                        let columnsList = searchObj.columns||[];
                        for(var i = 0; i < columnsList.length; i++){
                            var label = columnsList[i].label||'';
                            var name = columnsList[i].name||'';
                            if(label != ''){
                                if(name == 'formulacurrency' || name == 'formulanumeric' ){
                                    custList.addField({
                                        id: 'custpage_'+ (label.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                        type: ui.FieldType.FLOAT,
                                        label: label
                                    });
                                }else{
                                    custList.addField({
                                        id: 'custpage_'+ (label.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                        type: ui.FieldType.TEXT,
                                        label: label
                                    });
                                }
                            }else if(name != ''){
                                if(name == 'formulacurrency' || name == 'formulanumeric' ){
                                    custList.addField({
                                        id: 'custpage_'+ (name.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                        type: ui.FieldType.FLOAT,
                                        label: name
                                    });
                                }else{
                                    custList.addField({
                                        id: 'custpage_'+ (name.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                        type: ui.FieldType.TEXT,
                                        label: name
                                    });
                                }
                            }
                        }
                        //var resultSet = searchObj.runSearch();
                        var pagedSearch = searchObj.runPaged({
                            pageSize: DefaultListSize
                        });
                        var searchResultCount = pagedSearch.count;
                        form.updateDefaultValues({
                            custpage_totalcountfld : searchResultCount
                        });
                        var pageRangeList = pagedSearch.pageRanges;
                        //log.debug('pagedSearch.pageRanges' , pagedSearch.pageRanges);
                        //add page list options
                        var pageId = request.parameters.custpage_paginationfld||0;

                        //log.debug('searchResultCount', searchResultCount);
                        for(let p_i=0;p_i < pageRangeList.length; p_i++ ){
                            let value = pageRangeList[p_i].index;
                            let label = pageRangeList[p_i].compoundLabel;
                            paginationFld.addSelectOption({
                                value : value,
                                text : label
                            });
                        }

                        /* for(let p_i=1;p_i <= searchResultCount; p_i = (p_i+DefaultListSize) ){
                            let value = p_i+DefaultListSize > searchResultCount?searchResultCount:(p_i+DefaultListSize);
                            let label = p_i+' to '+(p_i+DefaultListSize > searchResultCount?searchResultCount:(p_i+DefaultListSize));
                            paginationFld.addSelectOption({
                                value : value,
                                text : label
                            });
                        }*/

                        if(searchResultCount > 0){
                            //Add Lists based on search result
                            try{
                                var myPage = pagedSearch.fetch({
                                    index: pageId
                                });
                                if (null != pageId && pageId != '') {
                                    //paginationFld.setDefaultValue(pageId);
                                    form.updateDefaultValues({
                                        custpage_paginationfld : pageId
                                    });
                                }
                            }catch(e){
                                var myPage = pagedSearch.fetch({
                                    index: 0
                                });
                                if (null != pageId && pageId != '') {
                                    //paginationFld.setDefaultValue(pageId);
                                    form.updateDefaultValues({
                                        custpage_paginationfld : ''
                                    });
                                }
                            }

                            let lineindex = 0;
                            myPage.data.forEach(function(result){

                                for(var v_i = 0; v_i < columnsList.length; v_i++){
                                    var label = columnsList[v_i].label||'';
                                    var name = columnsList[v_i].name||'';
                                    var valuetxt = result.getText(columnsList[v_i])||'';
                                    var value = result.getValue(columnsList[v_i])||'';
                                    var recId = result.id;
                                    //log.debug('valuetxt', valuetxt);

                                    if(label == 'VBN Date' || name == 'VBN Date'){
                                        log.debug('value', value);
                                        if(label != '' && VBNNotesDataList.hasOwnProperty(value) && VBNNotesDataList[value].vbnbillingdate != ''){
                                            log.debug('VBN Date true', 'vbn no :'+value + '|| VBNNotesDataList[value]'+VBNNotesDataList[value].vbnbillingdate);
                                            custList.setSublistValue({
                                                id: 'custpage_'+ (label.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                                line: lineindex,
                                                value: VBNNotesDataList[value].vbnbillingdate||''
                                            });
                                        }else if(name != '' && VBNNotesDataList.hasOwnProperty(value) && VBNNotesDataList[value].vbnbillingdate != ''){
                                            log.debug('VBN Date true', 'vbn no :'+value + '|| VBNNotesDataList[value]'+VBNNotesDataList[value].vbnbillingdate);
                                            custList.setSublistValue({
                                                id: 'custpage_'+ (name.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                                line: lineindex,
                                                value: VBNNotesDataList[value].vbnbillingdate||''
                                            });
                                        }
                                    }else{
                                        if(valuetxt != ''){
                                            if(name == 'formulacurrency' || name == 'formulanumeric' ){//check if value is string or number
                                                ///valuetxt = format.format({value:valuetxt, type: format.Type.FLOAT});
                                                if(valuetxt == 0){
                                                    valuetxt = '0.00'
                                                }
                                            }//if else then nothing to do
                                        }else if(value != ''){
                                            if(name == 'formulacurrency' || name == 'formulanumeric'){//check if value is string or number
                                                ///value = format.format({value:value, type: format.Type.FLOAT});
                                            }//if else then nothing to do
                                            if(value == 0){
                                                value = '0.00'
                                            }
                                        }
                                        if(valuetxt != ''){
                                            if(label != ''){
                                                custList.setSublistValue({
                                                    id: 'custpage_'+ (label.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                                    line: lineindex,
                                                    value: valuetxt
                                                });
                                            }else if(name != ''){
                                                custList.setSublistValue({
                                                    id: 'custpage_'+ (name.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                                    line: lineindex,
                                                    value: valuetxt
                                                });
                                            }
                                        }else if(value != ''){
                                            if(label != ''){
                                                custList.setSublistValue({
                                                    id: 'custpage_'+ (label.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                                    line: lineindex,
                                                    value: value
                                                });
                                            }else if(name != ''){
                                                custList.setSublistValue({
                                                    id: 'custpage_'+ (name.toLowerCase().replace(/[^a-z]/g, "")).substr(0,20).toLowerCase(),
                                                    line: lineindex,
                                                    value: value
                                                });
                                            }
                                        }
                                    }
                                }
                                lineindex++;
                            });
                            //add Submit Button
                            form.addSubmitButton({
                                label : 'Generate Excel Report'
                            });
                        }

                        //set Cleint script
                        form.clientScriptModulePath = 'SuiteScripts/JCurve/HandleAgingReportSuiteletEvent_CS.js';
                        //Write form object
                        scriptContext.response.writePage(form);
                        return;
                    }catch(e){
                        log.error("ERROR", e.toString());
                        scriptContext.response.write(e.toString());
                    }
                }else if(scriptContext.request.method == 'POST'){
                    var xmlString = '';
                    //Get filter values from parameters
                    var selectedAccount = request.parameters.custpage_accountflt||'';
                    var selectedFromDate = request.parameters.custpage_duedatefromflt||'';
                    var selectedToDate = request.parameters.custpage_duedatetoflt||'';
                    var selectedPaymentType = request.parameters.custpage_pymttypeflt||'';
                    var selectedPaymentBankName = request.parameters.custpage_pymtbnkname||'';
                    var selectedPaymentBankNameTxt = request.parameters.custpage_pymtbnknametxt||'';
                    var selectedPaymentTypeTxt = request.parameters.custpage_pymttypetxt||'';
                    var selectedDocType = request.parameters.custpage_doctypeflt||'';
                    var selectedDocTypeTxt = request.parameters.custpage_doctypetxt||'';
                    var selectedCurrency = request.parameters.custpage_currencyflt||'';
                    log.debug('request.parameters', request.parameters);
                    //add search Filters
                    /*if(selectedFromDate != ''){
                        searchObj.filters.push(search.createFilter({name: 'duedate', operator: search.Operator.ONORAFTER, values: selectedFromDate}));
                        summarySearchObj.filters.push(search.createFilter({name: 'duedate', operator: search.Operator.ONORAFTER, values: selectedFromDate}));
                    }
                    if(selectedToDate != ''){
                        searchObj.filters.push(search.createFilter({name: 'duedate', operator: search.Operator.ONORBEFORE, values: selectedToDate}));
                        summarySearchObj.filters.push(search.createFilter({name: 'duedate', operator: search.Operator.ONORBEFORE, values: selectedToDate}));
                    }*/
                    if(selectedFromDate != ''){
                        searchObj.filters.push(search.createFilter({
                            name: 'formuladate',
                            formula : "CASE {type} WHEN 'Bill' THEN {custbody_gisc_change_due}WHEN 'Bill Credit' THEN {custbody_gisc_change_due}WHEN 'Journal' THEN {trandate} ELSE {trandate}  END",
                            operator: search.Operator.ONORAFTER,
                            values: selectedFromDate
                        }));
                        summarySearchObj.filters.push(search.createFilter({
                            name: 'formuladate',
                            formula : "CASE {type} WHEN 'Bill' THEN {custbody_gisc_change_due}WHEN 'Bill Credit' THEN {custbody_gisc_change_due}WHEN 'Journal' THEN {trandate} ELSE {trandate}  END",
                            operator: search.Operator.ONORAFTER,
                            values: selectedFromDate
                        }));
                    }
                    if(selectedToDate != ''){
                        searchObj.filters.push(search.createFilter({
                            name: 'formuladate',
                            formula : "CASE {type} WHEN 'Bill' THEN {custbody_gisc_change_due}WHEN 'Bill Credit' THEN {custbody_gisc_change_due}WHEN 'Journal' THEN {trandate} ELSE {trandate}  END",
                            operator: search.Operator.ONORBEFORE,
                            values: selectedToDate
                        }));
                        summarySearchObj.filters.push(search.createFilter({
                            name: 'formuladate',
                            formula : "CASE {type} WHEN 'Bill' THEN {custbody_gisc_change_due}WHEN 'Bill Credit' THEN {custbody_gisc_change_due}WHEN 'Journal' THEN {trandate} ELSE {trandate}  END",
                            operator: search.Operator.ONORBEFORE,
                            values: selectedToDate
                        }));
                    }
                    if(selectedAccount != ''){
                        searchObj.filters.push(search.createFilter({name: 'account', operator: search.Operator.ANYOF, values: selectedAccount}));
                        summarySearchObj.filters.push(search.createFilter({name: 'account', operator: search.Operator.ANYOF, values: selectedAccount}));
                    }

                    if(selectedPaymentType != ''){
                        searchObj.filters.push(search.createFilter({name: 'custbody_vbn_paymentmethod', operator: search.Operator.ANYOF, values: selectedPaymentType.split(',')}));
                        summarySearchObj.filters.push(search.createFilter({name: 'custbody_vbn_paymentmethod', operator: search.Operator.ANYOF, values: selectedPaymentType.split(',')}));
                    }

                    if(selectedPaymentBankName != ''){
                        searchObj.filters.push(search.createFilter({name: 'custbody_bank_name', operator: search.Operator.ANYOF, values: selectedPaymentBankName.split(',')}));
                        summarySearchObj.filters.push(search.createFilter({name: 'custbody_bank_name', operator: search.Operator.ANYOF, values: selectedPaymentBankName.split(',')}));
                    }

                    if(selectedDocType != ''){
                        searchObj.filters.push(search.createFilter({name: 'custbody_sourcefrom', operator: search.Operator.ANYOF, values: selectedDocType.split(',')}));
                        summarySearchObj.filters.push(search.createFilter({name: 'custbody_sourcefrom', operator: search.Operator.ANYOF, values: selectedDocType.split(',')}));
                    }
                    if(selectedCurrency != ''){
                        searchObj.filters.push(search.createFilter({name: 'currency', operator: search.Operator.ANYOF, values: selectedCurrency.split(',')}));
                        summarySearchObj.filters.push(search.createFilter({name: 'currency', operator: search.Operator.ANYOF, values: selectedCurrency.split(',')}));
                    }
                    log.debug('today before converting', new Date());
                    var today = format.format({
                        value:new Date(),
                        type: format.Type.DATETIMETZ,
                        timezone: companyTimezone
                    });
                    log.debug('today after converting', today);
                    var formatedDateText = today;//today.getDate() + '/' + (today.getMonth()+1) + '/' + today.getFullYear() + ' ' + today.getHours() + ':' + today.getMinutes();
                    //Header Part
                    xmlString += excelHeader(selectedFromDate,selectedToDate, selectedPaymentType!=''?selectedPaymentTypeTxt.replace(/\u0005/g,','):'ALL', selectedDocType!=''?selectedDocTypeTxt.replace(/\u0005/g,','):'ALL', formatedDateText, searchObj.columns.length, selectedPaymentBankNameTxt!=''?selectedPaymentBankNameTxt.replace(/\u0005/g,','):'ALL');

                    //Column list
                    let columnsList = searchObj.columns||[];
                    xmlString += excelColumnLabel(columnsList);

                    //Row data
                    var pagedSearch = searchObj.runPaged({
                        pageSize: DefaultListSize
                    });
                    var searchResultCount = pagedSearch.count;
                    log.debug('searchResultCount' , searchResultCount);
                    var pageRangeList = pagedSearch.pageRanges;
                    for(let p_i=0;p_i < pageRangeList.length; p_i++ ){
                        let pageId = pageRangeList[p_i].index;

                        var myPage = pagedSearch.fetch({
                            index: pageId
                        });

                        myPage.data.forEach(function(result){
                            xmlString += excelRows(result, columnsList, VBNNotesDataList);
                        });
                    }

                    //Total of numeric column Row
                    columnsList = summarySearchObj.columns||[];
                    summarySearchObj.run().each(function(result){
                        xmlString += excelTotal(result, columnsList);
                        return false;
                    });


                    xmlString += excelFooter();

                    response.addHeader({
                        name: 'content-type',
                        value: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    });
                    var filename = "AgingPayableReport_"+(new Date().getTime())+".xls";
                    var fileObj = file.create({
                        name: filename,
                        fileType: file.Type.PLAINTEXT,
                        encoding: file.Encoding.UTF_8,
                        contents: xmlString
                    });
                    response.setHeader({ name: 'Content-Disposition', value: 'attachment; filename="' + filename + '"' });
                    response.writeFile(fileObj, true);
                }else{

                }
            }
            const excelHeader = (dueDate,selectedToDate, paymentType, docType, printReportDateTime, totalColumn, paymentBankName) => {
                var xmlString="";
                xmlString += "<?xml version=\"1.0\"?>";
                xmlString += "<?mso-application progid=\"Excel.Sheet\"?>";
                xmlString += "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"";
                xmlString += " xmlns:o=\"urn:schemas-microsoft-com:office:office\"";
                xmlString += " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"";
                xmlString += " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"";
                xmlString += " xmlns:html=\"http:\/\/www.w3.org\/TR\/REC-html40\">";
                xmlString += " <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">";
                xmlString += "  <Author>Microsoft Office User<\/Author>";
                xmlString += "  <LastAuthor>Microsoft Office User<\/LastAuthor>";
                xmlString += "  <Created>2022-09-08T09:00:59Z<\/Created>";
                xmlString += "  <LastSaved>2022-11-11T09:43:14Z<\/LastSaved>";
                xmlString += "  <Version>16.00<\/Version>";
                xmlString += " <\/DocumentProperties>";
                xmlString += " <OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\">";
                xmlString += "  <AllowPNG\/>";
                xmlString += " <\/OfficeDocumentSettings>";
                xmlString += " <ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">";
                xmlString += "  <WindowHeight>27100<\/WindowHeight>";
                xmlString += "  <WindowWidth>32767<\/WindowWidth>";
                xmlString += "  <WindowTopX>1580<\/WindowTopX>";
                xmlString += "  <WindowTopY>2000<\/WindowTopY>";
                xmlString += "  <ProtectStructure>False<\/ProtectStructure>";
                xmlString += "  <ProtectWindows>False<\/ProtectWindows>";
                xmlString += " <\/ExcelWorkbook>";
                xmlString += " <Styles>";
                xmlString += "  <Style ss:ID=\"Default\" ss:Name=\"Normal\">";
                xmlString += "   <Alignment ss:Vertical=\"Bottom\"\/>";
                xmlString += "   <Borders\/>";
                xmlString += "   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"12\" ss:Color=\"#000000\"\/>";
                xmlString += "   <Interior\/>";
                xmlString += "   <NumberFormat\/>";
                xmlString += "   <Protection\/>";
                xmlString += "  <\/Style>";
                xmlString += "  <Style ss:ID=\"s17\">";
                xmlString += "   <Font ss:FontName=\"Calibri\" x:CharSet=\"222\" x:Family=\"Swiss\" ss:Size=\"11\"";
                xmlString += "    ss:Color=\"#FF0000\"\/>";
                xmlString += "  <\/Style>";
                xmlString += "  <Style ss:ID=\"s21\">";
                xmlString += "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"\/>";
                xmlString += "  <\/Style>";
                xmlString += "  <Style ss:ID=\"s23\">";
                xmlString += "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"\/>";
                xmlString += "   <Borders>";
                xmlString += "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"\/>";
                xmlString += "    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"\/>";
                xmlString += "    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"\/>";
                xmlString += "    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"\/>";
                xmlString += "   <\/Borders>";
                xmlString += "  <\/Style>";
                xmlString += "  <Style ss:ID=\"s25\">";
                xmlString += "   <Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"\/>";
                xmlString += "  <\/Style>";
                xmlString += "  <Style ss:ID=\"s26\">";
                xmlString += "   <Borders>";
                xmlString += "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Double\" ss:Weight=\"3\"\/>";
                xmlString += "    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"\/>";
                xmlString += "   <\/Borders>";
                xmlString += "  <\/Style>";
                xmlString += "  <Style ss:ID=\"s27\">";
                xmlString += "   <Borders>";
                xmlString += "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Double\" ss:Weight=\"3\"\/>";
                xmlString += "    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"\/>";
                xmlString += "   <\/Borders>";
                xmlString += "   <NumberFormat ss:Format=\"Standard\"\/>";
                xmlString += "  <\/Style>";
                xmlString += "  <Style ss:ID=\"s28\">";
                xmlString += "   <Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"\/>";
                xmlString += "  <\/Style>";
                xmlString += "  <Style ss:ID=\"s68\">";
                xmlString += "      <NumberFormat ss:Format=\"Standard\"/>";
                xmlString += "  </Style>"
                xmlString += " <\/Styles>";
                xmlString += " <Worksheet ss:Name=\"Sheet1\">";
                xmlString += "  <Table x:FullColumns=\"1\"";
                xmlString += "   x:FullRows=\"1\" ss:DefaultColumnWidth=\"53\" ss:DefaultRowHeight=\"16\">";
                xmlString += "   <Column ss:Index=\"2\" ss:AutoFitWidth=\"0\" ss:Width=\"167\"\/>";
                xmlString += "   <Column ss:Width=\"201\"\/>";
                xmlString += "   <Column ss:Width=\"178\"\/>";
                xmlString += "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"76\"\/>";
                xmlString += "   <Column ss:Width=\"113\"\/>";
                xmlString += "   <Column ss:Width=\"176\"\/>";
                xmlString += "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"107\"\/>";
                xmlString += "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"75\"\/>";
                xmlString += "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"64\"\/>";
                xmlString += "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"75\"\/>";
                xmlString += "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"107\" ss:Span=\"1\"\/>";
                xmlString += "   <Column ss:Index=\"14\" ss:AutoFitWidth=\"0\" ss:Width=\"112\"\/>";
                xmlString += "   <Column ss:Width=\"102\"\/>";
                xmlString += "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"85\"\/>";
                xmlString += "   <Column ss:AutoFitWidth=\"0\" ss:Width=\"92\"\/>";
                xmlString += "   <Row>";
                xmlString += "    <Cell ss:Index=\""+ (totalColumn/2) +"\"><Data ss:Type=\"String\">GIS Co., Ltd.<\/Data><\/Cell>";
                xmlString += "    <Cell ss:Index=\""+ (totalColumn - 3) +"\" ss:MergeAcross=\"3\" ss:StyleID=\"s28\"><Data ss:Type=\"String\">Print Report Date Time : "+ printReportDateTime +"<\/Data><\/Cell>";
                xmlString += "    <Cell ss:StyleID=\"s17\"\/>";
                xmlString += "   <\/Row>";
                xmlString += "   <Row>";
                xmlString += "    <Cell ss:Index=\""+ (totalColumn/2) +"\"><Data ss:Type=\"String\">Aging Payable Report<\/Data><\/Cell>";
                xmlString += "    <Cell ss:Index=\""+ (totalColumn - 3) +"\" ss:MergeAcross=\"2\" ss:StyleID=\"s28\"><Data ss:Type=\"String\"><\/Data><\/Cell>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "   <\/Row>";
                xmlString += "   <Row>";
                xmlString += "    <Cell  ss:MergeAcross=\""+ totalColumn +"\" ss:StyleID=\"s21\"><Data ss:Type=\"String\">Due Date = "+ ((dueDate==''&&selectedToDate=='')?"ALL":"") + (dueDate!=''?dueDate:"") +""+ (selectedToDate!=''?(" To "+selectedToDate):"") +" ,ประเภทการจ่ายเงิน "+ paymentType +" ,ประเภทเอกสาร "+ docType +" ,ธนาคาร "+ paymentBankName +"<\/Data><\/Cell>";
                xmlString += "   <\/Row>";
                xmlString += "   <Row>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s21\"\/>";
                xmlString += "   <\/Row>";

                return xmlString;
            }

            const excelColumnLabel = (columnsList) => {
                var xmlString="";
                xmlString += "<Row>";
                for(var i = 0; i < columnsList.length; i++){
                    var label = columnsList[i].label||'';
                    var name = columnsList[i].name||'';
                    if(label != ''){
                        xmlString += "    <Cell ss:StyleID=\"s23\"><Data ss:Type=\"String\">"+ label +"<\/Data><\/Cell>";
                    }else if(name != ''){
                        xmlString += "    <Cell ss:StyleID=\"s23\"><Data ss:Type=\"String\">"+ name +"<\/Data><\/Cell>";
                    }
                }
                xmlString += "   <\/Row>";
                return xmlString;
            }

            const excelRows = (result, columnsList, VBNNotesDataList) => {
                var xmlString="";
                xmlString += "   <Row ss:Height=\"17\">";
                for(var v_i = 0; v_i < columnsList.length; v_i++){
                    var label = columnsList[v_i].label||'';
                    var name = columnsList[v_i].name||'';
                    var valuetxt = result.getText(columnsList[v_i])||'';
                    var value = result.getValue(columnsList[v_i])||'';
                    //log.debug('valuetxt', valuetxt);
                    //log.debug('value', value);
                    if(label == 'VBN Date' || name == 'VBN Date'){
                        log.debug('value', value);
                        log.debug("VBNNotesDataList.hasOwnProperty(value) && VBNNotesDataList[value].vbnbillingdate != ''", VBNNotesDataList.hasOwnProperty(value) && VBNNotesDataList[value].vbnbillingdate != '');
                        if(VBNNotesDataList.hasOwnProperty(value) && VBNNotesDataList[value].vbnbillingdate != ''){
                            xmlString += "    <Cell><Data ss:Type=\"String\">"+ VBNNotesDataList[value].vbnbillingdate +"<\/Data><\/Cell>";
                        }else{
                            xmlString += "    <Cell><Data ss:Type=\"String\"><\/Data><\/Cell>";
                        }
                    }else if(label == 'Exchange Rate' || name == 'Exchange Rate') {
                        if(valuetxt != ''){
                            /*if(valuetxt.match(/^\s*(\+|-)?(((\d+)(,)*)+(\.(\d*,*)*)?|(\.\d*))\s*$/)){//check if value is string or number
                                xmlString += "    <Cell ss:StyleID=\"s68\"><Data ss:Type=\"Number\">"+ valuetxt +"<\/Data><\/Cell>";
                            }else{
                                xmlString += "    <Cell><Data ss:Type=\"String\">"+ valuetxt +"<\/Data><\/Cell>";
                            }*/
                            if(name == 'formulacurrency' || name == 'formulanumeric' ){//check if value is string or number
                                xmlString += "    <Cell><Data ss:Type=\"Number\">"+ valuetxt +"<\/Data><\/Cell>";
                            }else{
                                xmlString += "    <Cell><Data ss:Type=\"String\">"+ valuetxt +"<\/Data><\/Cell>";
                            }
                        }else if(value != ''){
                            /*if(value.match(/^\s*(\+|-)?(((\d+)(,)*)+(\.(\d*,*)*)?|(\.\d*))\s*$/)){//check if value is string or number
                                xmlString += "    <Cell ss:StyleID=\"s68\"><Data ss:Type=\"Number\">"+ value +"<\/Data><\/Cell>";
                            }else{
                                xmlString += "    <Cell><Data ss:Type=\"String\">"+ value +"<\/Data><\/Cell>";
                            }*/
                            if(name == 'formulacurrency' || name == 'formulanumeric'){//check if value is string or number
                                xmlString += "    <Cell><Data ss:Type=\"Number\">"+ value +"<\/Data><\/Cell>";
                            }else{
                                xmlString += "    <Cell><Data ss:Type=\"String\">"+ value +"<\/Data><\/Cell>";
                            }
                        }else{
                            xmlString += "    <Cell ><Data ss:Type=\"String\"><\/Data><\/Cell>";
                        }
                    }else{
                        if(valuetxt != ''){
                            /*if(valuetxt.match(/^\s*(\+|-)?(((\d+)(,)*)+(\.(\d*,*)*)?|(\.\d*))\s*$/)){//check if value is string or number
                                xmlString += "    <Cell ss:StyleID=\"s68\"><Data ss:Type=\"Number\">"+ valuetxt +"<\/Data><\/Cell>";
                            }else{
                                xmlString += "    <Cell><Data ss:Type=\"String\">"+ valuetxt +"<\/Data><\/Cell>";
                            }*/
                            if(name == 'formulacurrency' || name == 'formulanumeric' ){//check if value is string or number
                                xmlString += "    <Cell ss:StyleID=\"s68\"><Data ss:Type=\"Number\">"+ valuetxt +"<\/Data><\/Cell>";
                            }else{
                                xmlString += "    <Cell><Data ss:Type=\"String\">"+ valuetxt +"<\/Data><\/Cell>";
                            }
                        }else if(value != ''){
                            /*if(value.match(/^\s*(\+|-)?(((\d+)(,)*)+(\.(\d*,*)*)?|(\.\d*))\s*$/)){//check if value is string or number
                                xmlString += "    <Cell ss:StyleID=\"s68\"><Data ss:Type=\"Number\">"+ value +"<\/Data><\/Cell>";
                            }else{
                                xmlString += "    <Cell><Data ss:Type=\"String\">"+ value +"<\/Data><\/Cell>";
                            }*/
                            if(name == 'formulacurrency' || name == 'formulanumeric'){//check if value is string or number
                                xmlString += "    <Cell ss:StyleID=\"s68\"><Data ss:Type=\"Number\">"+ value +"<\/Data><\/Cell>";
                            }else{
                                xmlString += "    <Cell><Data ss:Type=\"String\">"+ value +"<\/Data><\/Cell>";
                            }
                        }else{
                            xmlString += "    <Cell ><Data ss:Type=\"String\"><\/Data><\/Cell>";
                        }
                    }
                }

                xmlString += "   <\/Row>";
                return xmlString;
            }

            const excelTotal = (result, columnsList) => {
                var xmlString="";
                xmlString += "<Row ss:Height=\"17\">";

                for(var v_i = 0; v_i < columnsList.length; v_i++){
                    var label = columnsList[v_i].label||'';
                    var name = columnsList[v_i].name||'';
                    var summary = columnsList[v_i].summary||'';
                    var value = result.getValue(columnsList[v_i])||'';
                    log.debug('summary', summary);
                    log.debug('label', label);
                    log.debug('name', name);
                    log.debug('value', value);
                    log.debug('name.toLowerCase().indexOf("sum of")', name.toLowerCase().indexOf('sum of'));
                    log.debug('label.toLowerCase().indexOf("sum of")', label.toLowerCase().indexOf('sum of'));
                    if(label != ''){
                        if(value != '' && summary == 'SUM'){
                            xmlString += "     <Cell ss:StyleID=\"s27\"><Data ss:Type=\"Number\">"+ value +"<\/Data><\/Cell>";
                        }else if(label.toLowerCase() == 'total'){
                            xmlString += "    <Cell ss:StyleID=\"s26\"><Data ss:Type=\"String\">Total<\/Data><\/Cell>";
                        }else{
                            xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                        }
                    }else{
                        if(value != '' && summary == 'SUM'){
                            xmlString += "     <Cell ss:StyleID=\"s27\"><Data ss:Type=\"Number\">"+ value +"<\/Data><\/Cell>";
                        }else if(label.toLowerCase() == 'total'){
                            xmlString += "    <Cell ss:StyleID=\"s26\"><Data ss:Type=\"String\">Total<\/Data><\/Cell>";
                        }else{
                            xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                        }
                    }
                }
                /* xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"><Data ss:Type=\"String\">Total<\/Data><\/Cell>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s26\"\/>";
                xmlString += "    <Cell ss:StyleID=\"s27\"><Data ss:Type=\"Number\">999999999.99000001<\/Data><\/Cell>";
                xmlString += "    <Cell ss:StyleID=\"s27\"><Data ss:Type=\"Number\">99999.99<\/Data><\/Cell>";
                xmlString += "    <Cell ss:StyleID=\"s27\"><Data ss:Type=\"Number\">99999.99<\/Data><\/Cell>";
                xmlString += "    <Cell ss:StyleID=\"s27\"><Data ss:Type=\"Number\">999999999.99000001<\/Data><\/Cell>";*/
                xmlString += "   <\/Row>";
                return xmlString;
            }

            const excelFooter = () => {
                var xmlString="";
                //xmlString += "<Row ss:Height=\"17\"\/>";
                xmlString += "  <\/Table>";
                xmlString += "  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">";
                xmlString += "   <PageSetup>";
                xmlString += "    <Header x:Margin=\"0.3\"\/>";
                xmlString += "    <Footer x:Margin=\"0.3\"\/>";
                xmlString += "    <PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"\/>";
                xmlString += "   <\/PageSetup>";
                xmlString += "   <Selected\/>";
                xmlString += "   <Panes>";
                xmlString += "    <Pane>";
                xmlString += "     <Number>3<\/Number>";
                xmlString += "     <ActiveRow>13<\/ActiveRow>";
                xmlString += "     <ActiveCol>2<\/ActiveCol>";
                xmlString += "    <\/Pane>";
                xmlString += "   <\/Panes>";
                xmlString += "   <ProtectObjects>False<\/ProtectObjects>";
                xmlString += "   <ProtectScenarios>False<\/ProtectScenarios>";
                xmlString += "  <\/WorksheetOptions>";
                xmlString += " <\/Worksheet>";
                xmlString += "<\/Workbook>";
                xmlString += "";
                return xmlString;

            }

            return {onRequest}

        });
