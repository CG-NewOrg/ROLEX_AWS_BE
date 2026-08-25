import cds from '@sap/cds/eslint.config.mjs'

export default [
  ...cds.recommended,
  // Suppress console and unused-vars warnings only for specific files,
  // without changing application code or logic.
  {
    files: ['srv/cat-service.js'],
    rules: {
      'no-console': 'off',
      'no-unused-vars': 'off'
    }
  },
  {
    files: ['srv/server.js'],
    rules: {
      'no-console': 'off'
    }
  },
  {
    files: ['gen/**/*.js'],
    rules: {
      'no-console': 'off',
      'no-unused-vars': 'off'
    }
  }
]
