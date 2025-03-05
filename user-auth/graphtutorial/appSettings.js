const settings = {
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET_VALUE,
  tenantId: process.env.TENANT_ID,
  graphUserScopes: process.env.ENTRA_SCOPES,
};

module.exports = exports = {
  settings,
};
