const axios = require('axios');

/**
 * Calls the endpoint with authorization bearer token.
 * @param {string} endpoint
 * @param {string} accessToken
 */
async function callApi(endpoint, accessToken) {

    const options = {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    };

    console.log('request made to web API at: ' + new Date().toString());

    try {
        const response = await axios.default.get(endpoint, options);
        return response.data;
    } catch (error) {
        console.log(error)
        return error;
    }
};

async function getUsers(accessToken) {
    const endpoint = process.env.GRAPH_ENDPOINT + '/v1.0/users';
    const options = {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    };

    console.log('request made to web API at: ' + new Date().toString());
      try {
        const response = await axios.get(endpoint, options);
        return response.data;
    } catch (error) {
        console.log(error)
        return error;
    }
}

async function createTeam(accessToken) {
    const endpoint = process.env.GRAPH_ENDPOINT + '/v1.0/teams';
    const options = {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    };
    const reqData = {
        "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
        "displayName": "My Sample Team",
        "description": "My Sample Team’s Description"
      }
    console.log('request made to web API at: ' + new Date().toString());

    try {
        const response = await axios.post(endpoint, reqData, options);
        return response.data;
    } catch (error) {
        console.log(error)
        return error;
    }
}

async function getCalanders(accessToken) {
    const endpoint = process.env.GRAPH_ENDPOINT + '/v1.0/users/122c3ad4-6c13-49d3-9995-9b5b788c554c/calendars';
    const options = {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    };

    console.log('request made to web API at: ' + new Date().toString());
    const reqBody = {
        "name": "Volunteer"
      }
      try {
        const response = await axios.post(endpoint,reqBody, options);
        return response.data;
    } catch (error) {
        console.log(error)
        return error;
    }
}
module.exports = {
    callApi: callApi,
    getUsers: getUsers,
    createTeam: createTeam,
    getCalanders:getCalanders
};