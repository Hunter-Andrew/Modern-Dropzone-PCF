/** Unit tests for the control's logic. Run with `npm test`. */
module.exports = {
  testEnvironment: "jsdom",
  roots: ["<rootDir>/tests"],
  transform: {
    "^.+\\.tsx?$": [
      "ts-jest",
      { tsconfig: { jsx: "react", module: "commonjs", esModuleInterop: true, strict: true } },
    ],
  },
  moduleNameMapper: { "\\.css$": "<rootDir>/tests/styleStub.js" },
};
