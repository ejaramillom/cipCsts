require('dotenv').config();
const pino = require('pino');
const logger = pino({ level: 'info' });
const { keyInSelect } = require('readline-sync');
const {
  initializeGraphForAppOnlyAuth,
  getAppOnlyTokenAsync,
  getUsersAsync,
  makeGraphCallAsync,
} = require('./graphHelper.js');

async function main() {
  try {
    logger.info('DEBUG: 3. JavaScript Graph App-Only Tutorial');

    let choice = 0;
    const choices = ['Display access token', 'List users', 'Make a Graph call'];
    initializeGraphForAppOnlyAuth();
    choice = keyInSelect(choices, 'Select an option', { cancel: 'Exit' });
    switch (choice) {
      case -1:
        return logger.info(
          'DEBUG: 4. No choice was selected. Interrupting execution...',
        );
      case 0:
        // Display access token
        return await displayAccessTokenAsync();
      case 1:
        // List emails from user's inbox
        return await listUsersAsync();
      case 2:
        // Run any Graph code
        return await makeGraphCallAsync();
      default:
        return logger.info(
          `DEBUG: 5. Invalid choice. Please select one of the following: ${choices} Interrupting execution...`,
        );
    }
  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 6. unhandled error during main index execution. ',
    );

    throw error;
  }
}

async function displayAccessTokenAsync() {
  try {
    const appOnlyToken = await getAppOnlyTokenAsync();

    return logger.info(`DEBUG: 7. App-only token: ${appOnlyToken.slice(0, 5)}`);

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 8. unhandled error getting app-only access token. ',
    );

    throw error;
  }
}

async function listUsersAsync() {
  try {
    const userPage = await getUsersAsync();
    const users = userPage.value;
    for (const user of users) {
      logger.info(
        {
          user: `${user.displayName ?? 'NO NAME'}`,
          id: `${user.id ?? 'NO ID'}`,
          email: `${user.mail ?? 'NO EMAIL'}`,
        },
        `DEBUG: 9. User list of users: ${user}`,
      );
    }
    // Because if @odata.nextLink is not undefined, there are more users available on the server
    const moreAvailable = await moreUsersOnTheServer(userPage);

    return logger.warn(
      { usersInServer: moreAvailable },
      'DEBUG: 10. More users on the server? ',
    );

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 11. unhandled error getting app-only access token. ',
    );

    throw error;
  }
}

async function moreUsersOnTheServer(userPage) {
  try {

    return userPage['@odata.nextLink'] !== undefined;

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 12. Unhandled error while checking users on server. ',
    );

    throw error;
  }
}

(async () => {
  try {
    const result = await main();
    logger.info(
      { status: 200, message: result },
      'DEBUG: 1. App auth cip csts success! ',
    );

    process.exit(0);

  } catch (error) {
    logger.error(
      { name: error.name, message: error.message },
      'DEBUG: 2. App test failed: ',
    );

    throw error;
  }
})();
