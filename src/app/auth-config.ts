import { LogLevel, Configuration, BrowserCacheLocation } from '@azure/msal-browser';

export const msalConfig: Configuration = {
  auth: {
    clientId: '32751695-d57e-4fde-82a0-a8519d3708cf',
    authority: 'https://login.microsoftonline.com/1b88165d-e4e4-494e-bd92-bf0ae95fcdf1',
    redirectUri: 'http://localhost:4200',
    postLogoutRedirectUri: 'http://localhost:4200'
  },
  cache: {
    cacheLocation: BrowserCacheLocation.LocalStorage
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (!containsPii) console.log(message);
      },
      logLevel: LogLevel.Info
    }
  }
};

export const apiConfig = {
  uri: 'https://kc8cyrk24j.execute-api.us-east-1.amazonaws.com/api/pedidos',
  scopes: ['api://2ee2c2d7-159e-4c24-89dc-8d789e1de6a8/Pedidos.ReadWrite']
};