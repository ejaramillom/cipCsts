const globals = require('globals');
const js = require('@eslint/js');
const eslintPluginPrettier = require('eslint-plugin-prettier');
const eslintPrettierRecommended = require('eslint-plugin-prettier/recommended');
const header = require('eslint-plugin-header');

header.rules.header.meta.schema = false;

module.exports = [
  js.configs.recommended,
  eslintPrettierRecommended,
  {
    languageOptions: {
      globals: {
        ...globals.commonjs,
        ...globals.node,
      },

      ecmaVersion: 'latest',
      sourceType: 'module',
    },

    plugins: {
      header,
      eslintPluginPrettier,
    },

    rules: {
      'prettier/prettier': [
        'error',
        {
          singleQuote: true,
          endOfLine: 'auto',
          printWidth: 80,
        },
      ],
    },
  },
];
