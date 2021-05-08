// from https://docs.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0
// from https://developer.microsoft.com/en-us/graph/blogs/controlling-app-access-on-specific-sharepoint-site-collections/
// from https://ashiqf.com/2021/03/15/how-to-use-microsoft-graph-sharepoint-sites-selected-application-permission-in-a-azure-ad-application-for-more-granular-control/

// from https://stackoverflow.com/questions/19254029/angularjs-http-post-does-not-send-data
Object.toparams = function ObjecttoParams(obj) {
    var p = [];
    for (var key in obj) {
        p.push(key + '=' + encodeURIComponent(obj[key]));
    }
    return p.join('&');
};

function splistCtl($scope, $http) {
    // Defaults
    var vm = $scope;
    vm.hello = 'world';

    // M365 authentication
    // from https://code-boxx.com/simple-javascript-password-encryption-decryption/
    vm.getToken = function () {
        // from https://code-boxx.com/simple-javascript-password-encryption-decryption/
        // var original = "na6W0H1L.gyZcg-FJ~0gkY9N6-~42iVuHz";
        // console.log(original);
        // var secure = CryptoJS.AES.encrypt(original, document.location.host).toString();
        // console.log(secure);
        // var after = CryptoJS.AES.decrypt(secure, document.location.host).toString(CryptoJS.enc.Utf8);;
        // console.log(after);

        // JSON body
        var body = {
            "grant_type": "client_credentials",
            "client_id": "fe796649-2134-496e-8fd6-bbd90ff12a01",
            "client_secret": "T-wL2qRgRISqUm_B0xh.eZNo1n57U-RVmZ",
            "scope": "https://graph.microsoft.com/.default"
        }

        // HTTP POST
        var tenantName = "spjeff.com"
        var url = "https://login.microsoftonline.com/" + tenantName + "/oauth2/v2.0/token";
        $http({
            method: 'POST',
            url: url,
            data: Object.toparams(body),
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        }).then(function (resp) {
            // Parse
            vm.token = resp;
            $http.defaults.headers.common['Authorization'] = 'Bearer ' + resp.data.access_token;
        });
    }
    vm.getToken();

    // Add to SPList item
    vm.writeSPList = function () {
        // Config
        var siteId = "d4cad0de-2ea8-401a-bffc-bcac059af874"
        var listId = "7128f76b-b043-43b1-9a3c-060ba763e6e2"

        // JSON
        var body = {
            fields: {
                Title: 'hello world (selected.sites)'
            }
        };

        // POST
        $http({
            method: 'POST',
            url : 'https://graph.microsoft.com/v1.0/sites/' + siteId+ '/lists/' + listId+ '/items',
            data: body
        }).then(function (resp) {
            console.log(resp);
        });
    }
}


angular.module('splistApp', []).controller('splistCtl', splistCtl);