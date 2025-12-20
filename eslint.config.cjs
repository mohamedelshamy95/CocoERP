const appsScriptGlobals = {
  SpreadsheetApp: "readonly",
  ScriptApp: "readonly",
  PropertiesService: "readonly",
  LockService: "readonly",
  CacheService: "readonly",
  UrlFetchApp: "readonly",
  DriveApp: "readonly",
  Utilities: "readonly",
  Session: "readonly",
  HtmlService: "readonly",
  ContentService: "readonly",
  Logger: "readonly",
  console: "readonly"
};

module.exports = [
  {
    ignores: [
      "node_modules/**",
      ".vscode/**",
      ".eslint-bundle.js",
      "dist/**",
      "build/**"
    ]
  },
  {
    files: ["**/*.js"],
    languageOptions: {
      ecmaVersion: 2021,
      sourceType: "script",
      globals: appsScriptGlobals
    },
    rules: {
      "no-undef": "off",
      "no-unused-vars": ["warn", {
        "argsIgnorePattern": "^_",
        "caughtErrorsIgnorePattern": "^(?:_|e|e2|err|error)$"
      }],
      "no-restricted-properties": ["error",
        {
          "object": "SpreadsheetApp",
          "property": "getUi",
          "message": "Trigger-safe rule: never call SpreadsheetApp.getUi() from automation/trigger paths. UI allowed only in manual menu actions."
        }
      ]
    }
  },
  {
    files: ["AppCore.js"],
    rules: {
      "no-restricted-properties": "off"
    }
  }
];
