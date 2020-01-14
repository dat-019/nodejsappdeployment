require('dotenv').config();
var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');
var session = require('express-session');
var flash = require('connect-flash');
var passport = require('passport');
var OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
var graph = require('./graph');

// Configure simple-oauth2
const oauth2 = require('simple-oauth2').create({
    client: {
      id: process.env.OAUTH_APP_ID, //'38b723d2-0132-413f-aaf6-66027fed95e7',
      secret: process.env.OAUTH_APP_PASSWORD // .T]4efUWjK:Hi/0tOrdpThKyWxM8Q1?W
    },
    auth: {
      tokenHost: process.env.OAUTH_AUTHORITY, // 'https://login.microsoftonline.com/common',
      authorizePath: process.env.OAUTH_AUTHORIZE_ENDPOINT, // '/oauth2/v2.0/authorize'
      tokenPath: process.env.OAUTH_TOKEN_ENDPOINT // '/oauth2/v2.0/token'
    }
  });

// In-memory storage of logged-in users
// For demo purposes only
var users = {};

// Passport calls serializeUser and deserializeUser to
// manage users
passport.serializeUser(function(user, done) {
  // Use the OID property of the user as a key
  users[user.profile.oid] = user;
  done (null, user.profile.oid);
});

passport.deserializeUser(function(id, done) {
  done(null, users[id]);
});

// Callback function called once the sign-in is complete
// and an access token has been obtained
async function signInComplete(iss, sub, profile, accessToken, refreshToken, params, done) {
    if (!profile.oid) {
        console.log("No OID found in user profile");
      return done(new Error("No OID found in user profile."), null);
    }
  
    try{
      const user = await graph.getUserDetails(accessToken);
      console.log(user);
  
      if (user) {
        // Add properties to profile
        profile['email'] = user.mail ? user.mail : user.userPrincipalName;
      }
    } catch (err) {
        console.log(err);    
      done(err, null);
    }
  
    // Create a simple-oauth2 token from raw tokens
    let oauthToken = oauth2.accessToken.create(params);
  
    // Save the profile and tokens in user storage
    users[profile.oid] = { profile, oauthToken };
    return done(null, users[profile.oid]);
  }

  // Configure OIDC strategy
passport.use(new OIDCStrategy(
    {
      identityMetadata: `${process.env.OAUTH_AUTHORITY}${process.env.OAUTH_ID_METADATA}`, // 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
      clientID: process.env.OAUTH_APP_ID, // '38b723d2-0132-413f-aaf6-66027fed95e7'
      responseType: 'code id_token',
      responseMode: 'form_post',
      redirectUrl: process.env.OAUTH_REDIRECT_URI, // 'http://localhost:3000/auth/callback',
      allowHttpForRedirectUrl: true,
      clientSecret: process.env.OAUTH_APP_PASSWORD, // '.T]4efUWjK:Hi/0tOrdpThKyWxM8Q1?W',
      validateIssuer: false,
      passReqToCallback: false,
      loggingNoPII: false,
      scope: process.env.OAUTH_SCOPES.split(' ') // ['profile', '.default', 'offline_access']
      //OAUTH_SCOPES='profile offline_access Sites.Read.All Files.Read.All Files.ReadWrite.All Sites.ReadWrite.All'
    },
    signInComplete
  ));

var indexRouter = require('./routes/index');
var usersRouter = require('./routes/users');
var authRouter = require('./routes/auth');
var calendarRouter = require('./routes/calendar');
var tech6SPRouter = require('./routes/tech6');

var session = require('express-session');

var app = express();

app.use(function(req, res, next) {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
  next();
});

app.use(session({
    secret: 'poc_pgf!@#$%',
    resave: false,
    saveUninitialized: false,
    unset: 'destroy'
  }));

// Flash middleware
app.use(flash());
// Set up local vars for template layout
app.use(function(req, res, next) {
    // Read any flashed errors and save
    // in the response locals
    res.locals.error = req.flash('error_msg');
  
    // Check for simple error string and
    // convert to layout's expected format
    var errs = req.flash('error');
    for (var i in errs){
      res.locals.error.push({message: 'An error occurred', debug: errs[i]});
    }
  
    next();
  });

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// Initialize passport
app.use(passport.initialize());
app.use(passport.session());

app.use('/', indexRouter);
app.use('/users', usersRouter);
app.use('/auth', authRouter);
app.use('/calendar', calendarRouter);
app.use('/tech6', tech6SPRouter);

module.exports = app;
