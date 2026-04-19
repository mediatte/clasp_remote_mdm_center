import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');

const envPath = path.join(projectRoot, '.env');
const generatedConfigPath = path.join(projectRoot, 'Config.generated.js');
const claspPath = path.join(projectRoot, '.clasp.json');

const requiredKeys = ['APP_SCRIPT_ID', 'APP_SPREADSHEET_ID', 'APP_SENDER_EMAIL'];

main();

function main() {
  const env = parseEnvFile(readRequiredFile(envPath));
  const missingKeys = requiredKeys.filter((key) => !env[key]);

  if (missingKeys.length) {
    throw new Error(`Missing required .env keys: ${missingKeys.join(', ')}`);
  }

  writeFile(
    generatedConfigPath,
    buildGeneratedConfigSource({
      spreadsheetId: env.APP_SPREADSHEET_ID,
      senderEmail: env.APP_SENDER_EMAIL,
    })
  );

  writeFile(
    claspPath,
    JSON.stringify(
      {
        scriptId: env.APP_SCRIPT_ID,
        rootDir: '',
        scriptExtensions: ['.js', '.gs'],
        htmlExtensions: ['.html'],
        jsonExtensions: ['.json'],
        filePushOrder: ['Config.generated.js', 'Code.js'],
        skipSubdirectories: false,
      },
      null,
      2
    ) + '\n'
  );

  console.log('Generated Config.generated.js and .clasp.json from .env');
}

function readRequiredFile(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Missing ${path.basename(filePath)}. Copy .env.example to .env first.`);
  }

  return fs.readFileSync(filePath, 'utf8');
}

function parseEnvFile(source) {
  return source.split(/\r?\n/).reduce((accumulator, rawLine) => {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) {
      return accumulator;
    }

    const separatorIndex = line.indexOf('=');
    if (separatorIndex === -1) {
      return accumulator;
    }

    const key = line.slice(0, separatorIndex).trim();
    const value = stripWrappingQuotes(line.slice(separatorIndex + 1).trim());
    accumulator[key] = value;
    return accumulator;
  }, {});
}

function stripWrappingQuotes(value) {
  if (
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"))
  ) {
    return value.slice(1, -1);
  }

  return value;
}

function buildGeneratedConfigSource(config) {
  return `function getRuntimeSecretConfig_() {
  return Object.freeze({
    spreadsheetId: ${JSON.stringify(config.spreadsheetId)},
    senderEmail: ${JSON.stringify(config.senderEmail)},
  });
}
`;
}

function writeFile(filePath, contents) {
  fs.writeFileSync(filePath, contents, 'utf8');
}
