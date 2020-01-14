// DXC Tech6 items updates being maintained via SP list

const express = require('express');
const router = express.Router();
const tokens = require('../tokens.js');
const graph = require('../graph.js');
const moment = require('moment');
const nextLink = "@odata.nextLink";

router.get('/getTech6ItemList',
  async function (req, res, next) {
    var itemList = [];
    if (!req.isAuthenticated()) {
      res.redirect('/auth/signin'); // if the request is unauthenticated, redirect to login page
    }
    else {
      //get access token
      var accessToken;
      try {
        accessToken = await tokens.getAccessToken(req);
      } catch (err) {
        console.log("Could not get access token. Try signing out and signing in again.");
        console.log(JSON.stringify(err));
      }

      if (accessToken && accessToken.length > 0) {
        var requestUrl = "https://graph.microsoft.com/v1.0/sites/d7b730f3-e885-4dbe-934a-f37dc254977b/lists/4ed562b8-7e69-4257-81a7-abea7440b78c/items?$expand=fields";
        try {
          do {
            var list = await graph.getData(accessToken, requestUrl);
            if (list.value.length > 0) {
              (list.value).forEach(element => {
                itemList.push(element);
              });
            }
            if (list[nextLink]) {
              requestUrl = list[nextLink];
            }
          } while (list[nextLink]);
          console.log(itemList.length);
        } catch (error) {
          console.log(JSON.stringify(error));
          res.json(error);
        }
      }
      res.json(itemList);
    }
  }
);

// Get any Title of Item
router.get('/filterTech6ItemTitles/:title',
  async function (req, res, next) {
    var itemList = [];
    if (!req.isAuthenticated()) {
      res.redirect('/auth/signin'); // if the request is unauthenticated, redirect to login page
    }
    else {
      //get access token
      var accessToken;
      try {
        accessToken = await tokens.getAccessToken(req);
      } catch (err) {
        console.log("Could not get access token. Try signing out and signing in again.");
        console.log(JSON.stringify(err));
      }

      if (accessToken && accessToken.length > 0) {
        var filteredTitle = req.params.title;
        var requestUrl = "https://graph.microsoft.com/v1.0/sites/d7b730f3-e885-4dbe-934a-f37dc254977b/lists/4ed562b8-7e69-4257-81a7-abea7440b78c/items";
        try {
          do {
            var list = await graph.filterTech6ItemTitles(accessToken, requestUrl, filteredTitle);
            if (list.value.length > 0) {
              (list.value).forEach(element => {
                itemList.push(element);
              });
            }
            if (list[nextLink]) {
              requestUrl = list[nextLink];
            }
          } while (list[nextLink]);
          console.log(itemList.length);
        } catch (error) {
          console.log(JSON.stringify(error));
          res.json(error);
        }
      }
      res.json(itemList);
    }
  }
);

/* Get some items between two dates base on 'Event_x0020_Date'
  Input Date format: mm-dd-yyyy
  Ex: http://localhost:3000/tech6/filterTech6ItemEventDate/12-01-2018/01-01-2020
*/
router.get('/filterTech6ItemEventDate/:startEventDate/:endEventDate',
  async function (req, res, next) {
    var itemList = [];
    if (!req.isAuthenticated()) {
      res.redirect('/auth/signin'); // if the request is unauthenticated, redirect to login page
    }
    else {
      //get access token
      var accessToken;
      try {
        accessToken = await tokens.getAccessToken(req);
      } catch (err) {
        console.log("Could not get access token. Try signing out and signing in again.");
        console.log(JSON.stringify(err));
      }

      if (accessToken && accessToken.length > 0) {
        var p_StartEventDate = req.params.startEventDate ? new Date(req.params.startEventDate) : new Date();
        var p_EndEventDate = req.params.endEventDate ? new Date(req.params.endEventDate) : new Date();

        if (p_StartEventDate.getTime() > p_EndEventDate.getTime()) {
          res.send({ message: 'Start Event Date can not greater than End Event Date' });
        } else {
          var requestUrl = "https://graph.microsoft.com/v1.0/sites/d7b730f3-e885-4dbe-934a-f37dc254977b" +
            "/lists/4ed562b8-7e69-4257-81a7-abea7440b78c/items?$expand=fields&$filter=fields/Event_x0020_Date ge '" + p_StartEventDate.toISOString() + "'" +
            "and fields/Event_x0020_Date le '" + p_EndEventDate.toISOString() + "'";
          try {
            do {
              var list = await graph.filterTech6Item(accessToken, requestUrl);
              if (list.value.length > 0) {
                (list.value).forEach(element => {
                  itemList.push(element);
                });
              }
              if (list[nextLink]) {
                requestUrl = list[nextLink];
              }
            } while (list[nextLink]);
            console.log(itemList.length);
          } catch (error) {
            console.log(JSON.stringify(error));
            res.json(error);
          }
        }
      }
    }
    res.json(itemList);
  }
);

/* Get item with 'Event_x0020_Date' = specific date
  Input Date format: mm-dd-yyyy hh:mm:ss
  Ex: http://localhost:3000/tech6/filterTech6ItemEventDate/09-27-2018 24:00:00
*/
router.get('/filterTech6ItemEventDate/:specificEventDate',
  async function (req, res, next) {
    if (!req.isAuthenticated()) {
      res.redirect('/auth/signin'); // if the request is unauthenticated, redirect to login page
    }
    else {
      //get access token
      var accessToken;
      try {
        accessToken = await tokens.getAccessToken(req);
      } catch (err) {
        console.log("Could not get access token. Try signing out and signing in again.");
        console.log(JSON.stringify(err));
      }

      if (accessToken && accessToken.length > 0) {
        var p_SpecificEventDate = req.params.specificEventDate ? new Date(req.params.specificEventDate).toISOString() : new Date();

        var requestUrl = "https://graph.microsoft.com/v1.0/sites/d7b730f3-e885-4dbe-934a-f37dc254977b" +
          "/lists/4ed562b8-7e69-4257-81a7-abea7440b78c/items?$expand=fields&$filter=fields/Event_x0020_Date eq '" + p_SpecificEventDate + "'";
        try {

          var list = await graph.filterTech6Item(accessToken, requestUrl);
          res.json(list);

        } catch (error) {
          console.log(JSON.stringify(error));
          res.json(error);
        }
      }
    }
  }
);
module.exports = router;
