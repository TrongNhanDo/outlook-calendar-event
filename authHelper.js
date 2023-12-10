var clientId = '442cd972-613e-4273-84a6-f8e8a44ca867';
var clientSecret = '5Ak8Q~.lb33hzdcVGIhkL_quEbElA5VZW_w9db1w';
var redirectUri = 'http://localhost:3000/authorize';

var scopes = [
   'openid',
   'profile',
   'offline_access',
   'https://graph.microsoft.com/Calendars.ReadWrite',
];

var credentials = {
   clientID: clientId,
   clientSecret: clientSecret,
   site: 'https://login.microsoftonline.com/common',
   authorizationPath: '/oauth2/v2.0/authorize',
   tokenPath: '/oauth2/v2.0/token',
};
var oauth2 = require('simple-oauth2')(credentials);

module.exports = {
   getAuthUrl: function () {
      var returnVal = oauth2.authCode.authorizeURL({
         redirect_uri: redirectUri,
         scope: scopes.join(' '),
      });
      console.log({ generatedAuthUrl: returnVal });
      return returnVal;
   },

   getTokenFromCode: function (auth_code, callback, request, response) {
      oauth2.authCode.getToken(
         {
            code: auth_code,
            redirect_uri: redirectUri,
            scope: scopes.join(' '),
         },
         function (error, result) {
            if (error) {
               console.log({ accessTokenError: error.message });
               callback(request, response, error, null);
            } else {
               var token = oauth2.accessToken.create(result);
               console.log({ tokenCreated: token.token });
               callback(request, response, null, token);
            }
         }
      );
   },

   getEmailFromIdToken: function (id_token) {
      // JWT is in three parts, separated by a '.'
      var token_parts = id_token.split('.');

      // Token content is in the second part, in urlsafe base64
      var encoded_token = new Buffer(
         token_parts[1].replace('-', '+').replace('_', '/'),
         'base64'
      );

      var decoded_token = encoded_token.toString();

      var jwt = JSON.parse(decoded_token);

      // Email is in the preferred_username field
      return jwt.preferred_username;
   },

   getTokenFromRefreshToken: function (
      refresh_token,
      callback,
      request,
      response
   ) {
      var token = oauth2.accessToken.create({
         refresh_token: refresh_token,
         expires_in: 0,
      });
      token.refresh(function (error, result) {
         if (error) {
            console.log({ refreshTokenError: error.message });
            callback(request, response, error, null);
         } else {
            console.log({ newToken: result.token });
            callback(request, response, null, result);
         }
      });
   },
};
