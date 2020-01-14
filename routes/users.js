const express = require('express');
const router = express.Router();
const tokens = require('../tokens.js');
const graph = require('../graph.js');

const nextLink = "@odata.nextLink";
const value = "value";

// get current user detail
router.get('/getCurrentUserDetail',
  async function(req, res, next) {

    if (!req.isAuthenticated()) {
      res.redirect('/auth/signin'); // if the request is unauthenticated, redirect to login page
    }
    else
    {
        //get access token
        var accessToken;
        try
        {
          accessToken = await tokens.getAccessToken(req);
        } catch (err) {
          console.log("Could not get access token. Try signing out and signing in again.");
          console.log(JSON.stringify(err));
        }

        if (accessToken && accessToken.length > 0)
        {
            try {
              
              console.log(accessToken);
              var userDetail = await graph.getUserDetails(accessToken);
              res.json(userDetail);

            } catch (error) {
              console.log(JSON.stringify(error));
              res.json(error);
            }
        }
    }

  }
);

// Get all users without top parameter, by defautl the first 100 records will return
router.get('/getAllUsers', 
  async function(req, res, next) {

    if (!req.isAuthenticated()) {
      res.redirect('/auth/signin'); // if the request is unauthenticated, redirect to login page
    }
    else
    {
        //get access token
        var accessToken;
        try
        {
          accessToken = await tokens.getAccessToken(req);
        } catch (err) {
          console.log("Could not get access token. Try signing out and signing in again.");
          console.log(JSON.stringify(err));
        }

        if (accessToken && accessToken.length > 0)
        {
            try {
              console.log(accessToken);
              var users = await graph.getUsers(accessToken);
              res.json(users);

            } catch (error) {
              console.log(JSON.stringify(error));
              res.json(error);
            }
        }
    }
  }
);

// Get all users with top parameter is set
router.get('/getAllUsersV2', 
  async function(req, res, next) {
    var allUsers = [];
    if (!req.isAuthenticated()) {
      res.redirect('/auth/signin'); // if the request is unauthenticated, redirect to login page
    }
    else
    {
        var requestUrl = "https://graph.microsoft.com/v1.0/users?$top=999";
        while(requestUrl)
        {
            //get access token
            var accessToken;
            try
            {
              accessToken = await tokens.getAccessToken(req);
            } catch (err) {
              console.log("Could not get access token. Try signing out and signing in again.");
              console.log(JSON.stringify(err));
            }

            if (accessToken && accessToken.length > 0)
            {
                try {
                  
                  console.log(accessToken);
                  var users = await graph.getData(accessToken, requestUrl);
                  // res.json(users);
                  if (users)
                  {
                      usersVal = users[value];
                      if (allUsers.length === 0)
                      {
                          allUsers = usersVal;
                      }
                      else
                      {
                          var nextUserList = usersVal;
                          allUsers = allUsers.concat(nextUserList);
                      }
                      var nextLinkVal = users[nextLink];
                      if (nextLinkVal)
                      {
                        requestUrl = nextLinkVal;
                        // continue;
                      }
                      else
                      {
                        console.log("There is no next page.")  
                        break;
                      }
                  }

                } catch (error) {
                  console.log(JSON.stringify(error));
                  res.json(error);
                  return;
                }
            }
        } //end While block
    }
    res.json(allUsers);
  }
);

/* GET users with filter condition. */
router.get('/getFilteredUsers/:name', 
  async function(req, res, next) {

    if (!req.isAuthenticated()) {
      res.redirect('/auth/signin'); // if the request is unauthenticated, redirect to login page
    }
    else
    {
      var filteredName = req.params.name;
      // Search for users with the specifiec name across multiple properties.
      // var requestUrl = "https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'" + filteredName + "') or startswith(givenName,'" + filteredName + "') or startswith(surname,'" + filteredName + "') or startswith(mail,'" + filteredName + "') or startswith(userPrincipalName,'" + filteredName + "')&$top=999";
      var requestUrl = "https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'" + filteredName + "') or startswith(givenName,'" + filteredName + "') or startswith(surname,'" + filteredName + "') or startswith(mail,'" + filteredName + "') or startswith(userPrincipalName,'" + filteredName + "')";
      var accessToken;
      try
      {
        accessToken = await tokens.getAccessToken(req);
      } catch (err) {
        console.log("Could not get access token. Try signing out and signing in again.");
        console.log(JSON.stringify(err));
      }

      if (accessToken && accessToken.length > 0)
      {
          try {
            
            console.log(accessToken);
            var users = await graph.getData(accessToken, requestUrl);
            // res.json(users);
            if (users)
            {
              res.json(users);
            }
          } catch (error) {
            console.log(JSON.stringify(error));
            res.json(error);
          }
      }
    }
  }
);

module.exports = router;
