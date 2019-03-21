const Msal = require('msal');

const Event = function(){
  const handler = [];
  this.subscribe = function(callback){
    handler.push(callback);
  }
  this.unsubscribe = function(callback){
    handler = handler.filter(h => h !== callback);
  }
  this.emit = function(...data){
    handler.forEach(h => h(...data));
  }
}

class MsalWrapper {
  constructor(config) {
    const msalOptions = config.storeAuth !== false
      ? { storeAuthStateInCookie: true, cacheLocation: "localStorage" }
      : undefined;
    const msalClient = new Msal.UserAgentApplication(config.clientId, config.authority, acquireTokenRedirectCallback, msalOptions);
    let _idToken;
    // Browser check variables
    const ua = window.navigator.userAgent;
    const msie = ua.indexOf('MSIE ');
    const msie11 = ua.indexOf('Trident/');
    const msedge = ua.indexOf('Edge/');
    const isIE = msie > 0 || msie11 > 0;
    const isEdge = msedge > 0;
    this.onWarning = new Event();
    this.onError = new Event();
    this.onLogin = new Event();
    this.onLogout = new Event();
    this.login = function (scopes = config.scopes) {
      let user = msalClient.getUser();
      if (user) {
        onLogin.emit(user);
      }
      else {
        msalClient.loginPopup(scopes).then(function (idToken) {
          //Login Success
          _idToken = idToken;
          user = msalClient.getUser();
          onLogin.emit(user);
        }, function (error) {
          onError.emit(error);
        });
      }
    };
    this.logout = function () {
      onLogout.emit();
      msalClient.logout();
    };
    this.setXhrHeader = function (xhr, scopes = config.scopes) {
      acquireTokenSilent(scopes, accessToken => setRequestHeader(xhr, accessToken), error => onError.emit(error));
    };
    this.acquireToken = function (callback, scopes = config.scopes) {
      //Call acquireTokenSilent (iframe) to obtain a token for Microsoft Graph
      acquireTokenSilent(scopes, accessToken => callback(accessToken), error => onError.emit(error));
    };
    function setRequestHeader(xhr, accessToken) {
      xhr.setRequestHeader('Authorization', 'Bearer ' + accessToken);
      //xhr.send();
    }
    function acquireTokenSilent(scopes, completeCallback, errorCallback) {
      msalClient.acquireTokenSilent(scopes).then(function (accessToken) {
        if (completeCallback)
          completeCallback(accessToken);
      }, function (error) {
        acquireTokenSilentError(error, scopes, completeCallback, errorCallback);
      });
    }
    function acquireTokenPopup(scopes, completeCallback, errorCallback) {
      msalClient.acquireTokenPopup(scopes).then(function (accessToken) {
        if (completeCallback)
          completeCallback(accessToken);
      }, function (error) {
        if (errorCallback)
          errorCallback(error);
      });
    }
    function acquireTokenRedirect(scopes, completeCallback, errorCallback) {
      redirectCallbacks.push({ completeCallback, errorCallback });
      msalClient.acquireTokenRedirect(scopes);
    }
    function acquireTokenRedirectCallback(errorDesc, token, error, tokenType) {
      while (redirectCallbacks.length > 0) {
        const cb = redirectCallbacks.pop();
        if (error)
          cb.errorCallback ? cb.errorCallback(error) : onError.emit(error);
        else if (tokenType === "access_token") {
          if (cb.completeCallback)
            cb.completeCallback(token);
          else
            onWarning.emit('acquireTokenRedirect completed successfuly, but no callback was found!');
        }
        else {
          onError.emit(`token type is: "${tokenType}", not "access_token"`);
        }
      }
    }
    function acquireTokenSilentError(error, scopes, completeCallback, errorCallback) {
      // Call acquireTokenPopup (popup window) in case of acquireTokenSilent failure due to consent or interaction required ONLY
      if (error.indexOf("consent_required") !== -1
        || error.indexOf("interaction_required") !== -1
        || error.indexOf("login_required") !== -1) {
        onWarning.emit(error);
        if (isIE) {
          acquireTokenRedirect(scopes, completeCallback, errorCallback);
        }
        else {
          acquireTokenPopup(scopes, completeCallback, errorCallback);
        }
      }
      else {
        errorCallback(error);
      }
    }
  }
}

// allow usage in node & browser alike
if (typeof exports !== 'undefined') {
  if (typeof module !== 'undefined' && module.exports) {
    exports = module.exports = MsalWrapper;
  }
  exports.MsalWrapper = MsalWrapper;
} else if(typeof root !== 'undefined') {
  root['MsalWrapper'] = MsalWrapper;
}
