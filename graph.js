const graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const request = require('request');
const Q = require('q');


module.exports = {
  getUserDetails: async function (accessToken) {

    var client = getAuthenticatedClient(accessToken);

    var user = await client.api('/me').get();
    return user;
  },

  getEvents: async function (accessToken) {

    var client = getAuthenticatedClient(accessToken);

    var events = await client
      .api('/me/events')
      .select('subject,organizer,start,end')
      .orderby('createdDateTime DESC')
      .get();

    return events;
  },

  getUsers: async function (accessToken) {

    var client = getAuthenticatedClient(accessToken);
    var users = await client.api('/users?$top=999').get();
    return users;
  },


  getData: async function(accessToken, requestUrl) {

    var client = getAuthenticatedClient(accessToken);
    var responsedBlog = await client.api(requestUrl).get();
    return responsedBlog;
  },

  filterTech6ItemTitles: async function (accessToken, requestUrl, filteredTitle) {
    var client = getAuthenticatedClient(accessToken);
    var responsedBlog = await client
      .api(requestUrl)
      .select("id", "createdBy", "fields")
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .expand("fields")
      .filter("startswith(fields/Title, '" + filteredTitle + "') ")
      .get();
    return responsedBlog;
  },

  filterTech6Item: async function (accessToken, requestUrl) {
    var client = getAuthenticatedClient(accessToken);
    var responsedBlog = await client
      .api(requestUrl)
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .top(999)
      .get();
    return responsedBlog;
  },

  getDataV2: function(accessToken, requestUrl) {

      var deferred = Q.defer();
      request.get(requestUrl, {
        auth: {
            bearer: accessToken
          }
        }, 
        function (err, response, body) {
          var parsedBody = JSON.parse(body);

          if (err) {
            deferred.reject(err);
          } else if (parsedBody.error) {
            deferred.reject(parsedBody.error.message);
          } else {
            // The value of the body
            deferred.resolve(parsedBody.value);
            console.log(parsedBody);
          }
        }
    );
    return deferred.promise;
  }
};

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    defaultVersion: "v1.0",
    debugLogging: true,
    // Use the provided access token to authenticate requests
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  return client;
}