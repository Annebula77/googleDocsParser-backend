module.exports = {
  root: true,
  env: {
    es2020: true,
    node: true, //окружение node вместо browser
  },
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
    'plugin:prettier/recommended',
  ],
  ignorePatterns: ['dist', '.eslintrc.js'],
  parser: '@typescript-eslint/parser',
  plugins: [
    '@typescript-eslint',
    'prettier'
  ],
  rules: {
    'prettier/prettier': 'error',
  },
  settings: {
    'import/resolver': {
      typescript: {}, // Это гарантирует, что ESLint будет разрешать модули с использованием конфигурации tsconfig.json
    },
  },
};
