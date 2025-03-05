require('dotenv').config();
const pino = require('pino');
const logger = pino({ level: 'info' });
require('isomorphic-fetch');
const { DeviceCodeCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const {
  TokenCredentialAuthenticationProvider,
} = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js');

let _settings;
let _deviceCodeCredential;
let _userClient;

function initializeGraphForUserAuth(settings, deviceCodePrompt) {
  if (!settings) {
    throw new Error('DEBUG: 13. Settings cannot be undefined');
  }

  _settings = settings;

  _deviceCodeCredential = new DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt,
  });

  const authProvider = new TokenCredentialAuthenticationProvider(
    _deviceCodeCredential,
    {
      scopes: `${settings.graphUserScopes}`.split(','),
    },
  );

  _userClient = Client.initWithMiddleware({
    authProvider: authProvider,
  });
}

async function getUserTokenAsync() {
  // Ensure credential isn't undefined
  if (!_deviceCodeCredential) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Ensure scopes isn't undefined
  if (!_settings?.graphUserScopes) {
    throw new Error('Setting "scopes" cannot be undefined');
  }

  // Request token with given scopes
  const response = await _deviceCodeCredential.getToken(
    _settings?.graphUserScopes,
  );
  return response.token;
}

async function getUserAsync() {
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Only request specific properties with .select()
  return _userClient
    .api('/me')
    .select(['displayName', 'mail', 'userPrincipalName'])
    .get();
}

async function getInboxAsync() {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient
    .api('/me/mailFolders/inbox/messages')
    .select(['from', 'isRead', 'receivedDateTime', 'subject'])
    .top(25)
    .orderby('receivedDateTime DESC')
    .get();
}

async function sendMailAsync(subject, body, recipient) {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Create a new message
  const message = {
    subject: subject,
    body: {
      content: body,
      contentType: 'text',
    },
    toRecipients: [
      {
        emailAddress: {
          address: recipient,
        },
      },
    ],
  };

  // Send the message
  return _userClient.api('me/sendMail').post({
    message: message,
  });
}

// This function serves as a playground for testing Graph snippets
// or other code
async function makeGraphCallAsync() {
  // INSERT YOUR CODE HERE
  try {
    return logger.info('somethings');

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 12. Unhandled error while making graph call async. ',
    );

    throw error;
  }
}

module.exports = {
  makeGraphCallAsync,
  initializeGraphForUserAuth,
  getUserTokenAsync,
  sendMailAsync,
  getUserAsync,
  getInboxAsync,
};
