import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';

export function generateMcpTools(openapiTrimmedFile, clientFilePath) {
  try {
    console.log(
      `Generating ${path.basename(clientFilePath)} from ${path.basename(openapiTrimmedFile)} using openapi-zod-client...`
    );

    const outputDir = path.dirname(clientFilePath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`Created directory: ${outputDir}`);
    }

    // Pinned: unpinned `npx -y openapi-zod-client` resolves "latest" independently
    // on every host/CI runner/Docker build, so the tool's output shape (and thus
    // whether the @zodios/core rewrite below still matches) can drift between two
    // otherwise-identical `npm run generate` invocations. 1.18.3 is the version
    // that was actually resolved (and verified working end-to-end: npm run verify
    // green, image smoke-tested) as of this pin — bump deliberately, not by drift.
    execSync(
      `npx -y openapi-zod-client@1.18.3 "${openapiTrimmedFile}" -o "${clientFilePath}" --with-description --strict-objects --additional-props-default-value=false`,
      {
        stdio: 'inherit',
      }
    );

    console.log(`Generated client code at: ${clientFilePath}`);

    let clientCode = fs.readFileSync(clientFilePath, 'utf-8');
    clientCode = clientCode.replace(/'@zodios\/core';/, "'./hack.js';");

    clientCode = clientCode.replace(/\.strict\(\)/g, '.passthrough()');

    console.log('Stripping unused errors arrays from endpoint definitions...');
    // I didn't make up this crazy regex myself; you know who did. It seems works though.
    clientCode = clientCode.replace(/,?\s*errors:\s*\[[\s\S]*?],?(?=\s*})/g, '');

    console.log('Decoding HTML entities in path patterns...');
    // openapi-zod-client HTML-encodes special characters in path patterns
    // This breaks Microsoft Graph function-style APIs like range(address='A1:G10')
    clientCode = clientCode.replace(/&#x3D;/g, '='); // Decode = sign
    clientCode = clientCode.replace(/&#x27;/g, "'"); // Decode single quote
    clientCode = clientCode.replace(/&#x28;/g, '('); // Decode left paren
    clientCode = clientCode.replace(/&#x29;/g, ')'); // Decode right paren
    clientCode = clientCode.replace(/&#x3A;/g, ':'); // Decode colon

    console.log('Fixing function-style API paths with template literals...');
    // After HTML decoding, paths like range(address=':address') have nested single quotes
    // which cause TypeScript syntax errors. Convert the path string from single quotes
    // to backticks (template literal) so single quotes can remain inside.
    // Match: path: '/...range(param=':value')...',
    // Replace with: path: `/...range(param=':value')...`,
    clientCode = clientCode.replace(/(path:\s*)'(\/[^']*\([^)]*=':[\w]+'\)[^']*)'/g, '$1`$2`');

    // openapi-zod-client emits z.instanceof(File) for `format: binary` bodies; MCP
    // transports JSON so no caller produces File. Body marshaller decodes the string.
    clientCode = clientCode.replace(
      /z\.instanceof\(File\)/g,
      "z.string().describe('Base64-encoded file content. The server decodes it and PUTs the raw bytes to Microsoft Graph.')"
    );

    fs.writeFileSync(clientFilePath, clientCode);

    // Format the generated client so `npm run generate` output is prettier-stable and
    // the format:check step in `npm run verify` passes deterministically across versions.
    console.log('Formatting generated client with Prettier...');
    execSync(`npx prettier --write "${clientFilePath}"`, { stdio: 'inherit' });

    return true;
  } catch (error) {
    throw new Error(`Error generating client code: ${error.message}`);
  }
}
