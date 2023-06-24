/*

  handle dataverse/d365 authentification + authorization

*/

var msal = require('@azure/msal-node');
var axios = require('axios');

const msClientId = process.env.DS365_CLIENT_ID;
const msTenantId = process.env.D365_TENANT_ID;
const msClientSecret = process.env.D365_CLIENT_SECRET;
const DynamicsBaseUrl = process.env.D365_BASE_URL;
const DynamicsAPIPath = 'api/data/v9.2/';

var auth = {

  sayHi: function() {
    console.log('hi');
    return 'hi';
  },

  generateAccessToken: async function () {

    try {
      const msalConfig = {
        auth: {
          clientId: msClientId,
          clientSecret: msClientSecret,
          authority: `https://login.microsoftonline.com/${msTenantId}`,
        }
      }
      const cca = new msal.ConfidentialClientApplication(msalConfig);
      const authResponse = await cca.acquireTokenByClientCredential({
        scopes: [`${DynamicsBaseUrl}.default`]
      });

      return authResponse.accessToken;
    } catch (e) {
      console.error(e.message);
      return e.message;
    }

  },

  performAPICall: async function (method, path, http_params) {
    
    const url = `${DynamicsBaseUrl}${DynamicsAPIPath}${path}`;

    var http_params = http_params || {};
    http_params = JSON.stringify(http_params);

    console.debug(`D365 ${method} calling ${url} `);

    let token = await this.generateAccessToken();

    let headers = {
      'OData-MaxVersion': '4.0', 
      'Authorization': `Bearer ${token}`,
      'Prefer': 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
      'Content-Type': 'application/json'
    };

    var config = {
      method: method,
      url: url,
      headers: headers,
      data: http_params
    };

    let call = await axios(config);
    let data = call.data.value;

    return data;

  }

}

module.exports = auth