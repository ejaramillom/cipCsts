require('dotenv').config();
require('isomorphic-fetch');
const pino = require('pino');
const logger = pino({ level: 'info' });
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const {
  TokenCredentialAuthenticationProvider,
} = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js');

let _settings;
let _clientSecretCredential;
let _appClient;

function initializeGraphForAppOnlyAuth() {
  try {
    _settings = {
      tenantId: process.env.TENANT_ID,
      clientId: process.env.CLIENT_ID,
      clientSecret: process.env.CLIENT_SECRET_VALUE,
    };
    checkEmptySettings(_settings);

    if (!_clientSecretCredential) {
      _clientSecretCredential = new ClientSecretCredential(
        _settings.tenantId,
        _settings.clientId,
        _settings.clientSecret,
      );
    }

    if (!_appClient) {
      const authProvider = new TokenCredentialAuthenticationProvider(
        _clientSecretCredential,
        {
          scopes: ['https://graph.microsoft.com/.default'],
        },
      );

      _appClient = Client.initWithMiddleware({
        authProvider: authProvider,
      });
    }
  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 1. Unhandled error during authentication. ',
    );

    throw error;
  }
}

async function getAppOnlyTokenAsync() {
  try {
    checkEmptyCredentials(_clientSecretCredential);

    const response = await _clientSecretCredential.getToken([
      'https://graph.microsoft.com/.default',
    ]);

    return response.token;
  } catch (error) {
    logger.error(
      {
        name: error.name,
        message: error.message,
      },
      'DEBUG: 2. Unhandled error get app only token. ',
    );

    throw error;
  }
}

async function getUsersAsync() {
  try {
    checkEmptyAppClient(_appClient);
    // logger.info({ appClient: _appClient }, 'DEBUG: 999. check empty client');

    return _appClient
      ?.api('/users')
      .select(['displayName', 'id', 'mail'])
      .top(3)
      .orderby('displayName')
      .get();
  } catch (error) {
    logger.error(
      {
        name: error.name,
        message: error.message,
      },
      'DEBUG: 3. Unhandled error in getting user. ',
    );

    throw error;
  }
}

// This function serves as a playground for testing Graph snippets or other code
async function makeGraphCallAsync() {
  try {
    return logger.info(`DEBUG: 4. Graph call called`);
    // INSERT YOUR CODE HERE
  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 5. Unhandled error in graph call in helper. ',
    );

    throw error;
  }
}

function checkEmptySettings(settings) {
  if (!settings) {
    logger.error(
      {
        name: 'Undefined settings',
        message: 'Settings cannot be undefined',
      },
      'DEBUG: 6. unhandled error with undefined settings during graph initialization for app auth',
    );

    throw new Error('Settings cannot be undefined');
  }
}

function checkEmptyCredentials(clientSecretCredential) {
  if (!clientSecretCredential) {
    logger.error(
      {
        name: 'Uninitialized graph',
        message:
          'Graph credentials have not been initialized for app-only auth',
      },
      'DEBUG: 7. Unhandled error during client secret credential generation. ',
    );

    throw new Error(
      'Graph credentials have not been initialized for app-only auth',
    );
  }
}

function checkEmptyAppClient(appClient) {
  if (!appClient) {
    logger.error(
      {
        name: 'Uninitialized graph',
        message: 'Graph app client has not been initialized for app-only auth',
      },
      'DEBUG: 8. Unhandled error during client secret credential generation. ',
    );

    throw new Error(
      'Graph app client has not been initialized for app-only auth',
    );
  }
}

module.exports = {
  makeGraphCallAsync,
  getAppOnlyTokenAsync,
  getUsersAsync,
  initializeGraphForAppOnlyAuth,
};
