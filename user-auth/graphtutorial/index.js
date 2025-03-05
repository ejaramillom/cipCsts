require('dotenv').config();
const pino = require('pino');
const logger = pino({ level: 'info' });
const { keyInSelect } = require('readline-sync');
const {
  initializeGraphForUserAuth,
  getUserAsync,
  getUserTokenAsync,
  getInboxAsync,
  sendMailAsync,
  makeGraphCallAsync,
} = require('./graphHelper.js');
const settings = require('./appSettings.js').settings;

async function main() {
  try {
    let choice = 0;
    initializeGraph(settings);
    await greetUserAsync();
    const choices = [
      'Display access token',
      'List my inbox',
      'Send mail',
      'Make a Graph call',
    ];

    choice = keyInSelect(choices, 'Select an option', { cancel: 'Exit' });

    switch (choice) {
      case -1:
        return logger.info(
          'DEBUG: 3. No choice was selected. Interrupting execution...',
        );
      case 0:
        // Display access token
        return await displayAccessTokenAsync();
      case 1:
        // List emails from user's inbox
        return await listInboxAsync();
      case 2:
        // Send an email message
        return await sendMailToSelfAsync();
      case 3:
        // Run any Graph code
        return await makeGraphCallAsync();
      default:
        return logger.info(
          `DEBUG: 4. Invalid choice. Please select one of the following: ${choices} Interrupting execution...`,
        );
    }
  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 5. unhandled error during main user index execution. ',
    );

    throw error;
  }
}

function initializeGraph(settings) {
  try {
    initializeGraphForUserAuth(settings, (info) => {
      // Display the device code message to
      // the user. This tells them
      // where to go to sign in and provides the
      // code to use.
      logger.info(
        { info: info.message },
        'DEBUG: 6. Message from device login',
      );
    });
  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 7. unhandled error during user authentication. ',
    );

    throw error;
  }
}

async function greetUserAsync() {
  try {
    const user = await getUserAsync();
    logger.info(`Hello, ${user?.displayName}!`);
    // For Work/school accounts, email is in mail property
    // Personal accounts, email is in userPrincipalName
    logger.info(`Email: ${user?.mail ?? user?.userPrincipalName ?? ''}`);

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 8. unhandled error while greeting user. ',
    );

    throw error;
  }
}
async function displayAccessTokenAsync() {
  try {
    const userToken = await getUserTokenAsync();

    logger.info({ token: userToken.slice(0, 5) }, 'DEBUG: 9. User token.');

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 10. unhandled error while greeting user. ',
    );

    throw error;
  }
}
async function listInboxAsync() {
  try {
    const messagePage = await getInboxAsync();
    const messages = messagePage.value;

    for (const message of messages) {
      logger.info(`Message: ${message.subject ?? 'NO SUBJECT'}`);
      logger.info(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
      logger.info(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
      logger.info(`  Received: ${message.receivedDateTime}`);
    }

    // If @odata.nextLink is not undefined, there are more messages
    // available on the server
    const moreAvailable = messagePage['@odata.nextLink'] !== undefined;
    logger.info(`More messages available? ${moreAvailable}`);

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 11. unhandled error while listing user inbox. ',
    );

    throw error;
  }
}
async function sendMailToSelfAsync() {
  try {
    // Send mail to the signed-in user
    // Get the user for their email address
    const user = await getUserAsync();
    const userEmail = user?.mail ?? user?.userPrincipalName;
    if (!userEmail) {
      return logger.info("Couldn't get your email address, canceling...");
    }
    await sendMailAsync('Testing Microsoft Graph', 'Hello world!', userEmail);
    logger.info('Mail sent.');

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 12. unhandled error while sending user email. ',
    );

    throw error;
  }
}

(async () => {
  try {
    const result = await main();
    logger.info(
      { status: 200, message: result },
      'DEBUG: 1.1 User auth cip csts success! ',
    );

    process.exit(0);

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 2.1 User Test failed: ',
    );

    throw error;
  }
})();
