var express = require('express');
var passport = require('passport');
var router = express.Router();

/* GET auth callback. */
router.get('/signin',
  function(req, res, next) {
    passport.authenticate('azuread-openidconnect',
    { 
        response: res,
        prompt: 'login', // 'consent',
        failureRedirect: '/',
        failureFlash: true,
        successRedirect: '/'
    }
    )(req, res, next);

  }
);

router.post('/callback',
  function(req, res, next) {
    passport.authenticate('azuread-openidconnect',
    {
      response: res,
      failureRedirect: '/failure',
      failureFlash: true,
      successRedirect: '/' //the redirect url will be called upon calling back suceessfully
      // successRedirect: 'http://localhost:4200/'
    }
    )(req, res, next);
  }
);

router.get('/signout',
  function(req, res) {
    req.session.destroy(function(err) {
      req.logout();
      // res.redirect('/');
      res.redirect('http://localhost:4200/');
    });
  }
);

router.get('/isAuthenticated',
  function (req, res, next)
  {
    if (req.isAuthenticated()) {
        res.json(true);
    }
    else {
      res.json(false);
    }
  }
);

module.exports = router;