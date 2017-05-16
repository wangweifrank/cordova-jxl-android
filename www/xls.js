    var xls =  {
        save: function(data, dirname, filename, sheetdata, successCallback, errorCallback){
            cordova.exec(
                successCallback,// success callback
                errorCallback,  // error callback
                'Xls',          // class
                'saveXLS',      // action
                [{              // Array of param
                    "data"     : data,
                    "dirname"  : dirname,
                    "filename" : filename,
                    "sheetdata": sheetdata
                }]
            );
        }
    }

    module.exports = xls;
