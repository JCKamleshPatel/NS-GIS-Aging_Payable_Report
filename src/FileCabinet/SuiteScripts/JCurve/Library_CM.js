/**
* @NApiVersion 2.1
*/
/*
* @name:                                       Library_CM.js
* @author:                                     Kamlesh Patel
* @summary:                                    This script contain common & reusable library functions. 
* @copyright:                                  © Copyright by Jcurve Solutions
* Date Created:                                Sat Aug 24 2022 1:32:58 PM
* Change Logs:
* Date                          Author               Description
* Sat Aug 24 2022 1:32:58 PM -- Kamlesh Patel -- Initial Creation
*/

define(['N/format', 'N/runtime', 'N/search'],

(format, runtime, search) => {
    
    /**
    * runSearch : runs the search (s) and returns array of searchResults
    * @param {search} search
    * @return {SearchResultSet} searchResultSet
    */
    const runSearch = (search) => {
        
        var searchResults = [];
        try {
            if (!search) return [];
            var results = search.run(), searchid = 0;
            do {
                var resultslice = results.getRange({start:searchid,end:searchid+1000});
                resultslice.forEach(function(slice) {
                    searchResults.push(slice);
                    searchid++;
                }
                );
            } while (resultslice.length >=1000);
        } catch (e) {log.error(e.name,e.message+(e.stack?' | '+e.stack:'') )}
        return searchResults;
    }
    
    return {runSearch}
    
});
