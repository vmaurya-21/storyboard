module.exports = {
    preset: 'ts-jest',
    testEnvironment: 'jsdom',
    roots: ['<rootDir>/src'],
    testMatch: ['**/__tests__/**/*.test.ts?(x)'],
    moduleNameMapper: {
        // Mock CSS/SCSS modules
        '\\.(css|scss)$': 'identity-obj-proxy',

        // Mock SPFx modules
        '^@microsoft/sp-core-library$': '<rootDir>/src/__mocks__/@microsoft/sp-core-library.ts',
        '^@microsoft/sp-webpart-base$': '<rootDir>/src/__mocks__/@microsoft/sp-webpart-base.ts',
        '^@microsoft/sp-component-base$': '<rootDir>/src/__mocks__/@microsoft/sp-component-base.ts',
        '^@microsoft/sp-property-pane$': '<rootDir>/src/__mocks__/@microsoft/sp-property-pane.ts',
        '^@pnp/sp(|/.*)$': '<rootDir>/src/__mocks__/@pnp/sp.ts',
        '^@pnp/core(|/.*)$': '<rootDir>/src/__mocks__/@pnp/core.ts',

        // Mock localization strings
        '^SpxReporterDailyStoryDragDropWebPartStrings$': '<rootDir>/src/__mocks__/SpxReporterDailyStoryDragDropWebPartStrings.ts',

        // Mock Fluent UI
        '^@fluentui/react/lib/Icon$': '<rootDir>/src/__mocks__/@fluentui/react-icon.ts',
    },
    setupFilesAfterEnv: ['<rootDir>/src/setupTests.ts'],
    collectCoverageFrom: [
        'src/**/*.{ts,tsx}',
        '!src/**/*.d.ts',
        '!src/**/*.module.scss.ts',
        '!src/**/index.ts'
    ],
    coverageThreshold: {
        global: {
            branches: 50,
            functions: 50,
            lines: 50,
            statements: 50
        }
    },
    transform: {
        '^.+\\.[jt]sx?$': ['ts-jest', {
            tsconfig: {
                jsx: 'react',
                esModuleInterop: true,
                allowJs: true
            }
        }]
    },
    transformIgnorePatterns: [
        'node_modules/(?!(@dnd-kit|@fluentui|@microsoft/sp-lodash-subset|@pnp)/)'
    ]
};
