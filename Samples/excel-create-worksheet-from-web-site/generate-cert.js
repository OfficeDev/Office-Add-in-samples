// Script to generate self-signed certificates for local HTTPS development
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

const certPath = path.join(__dirname, 'localhost.crt');
const keyPath = path.join(__dirname, 'localhost.key');

// Check if certificates already exist
if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    console.log('SSL certificates already exist!');
    console.log('  - localhost.crt');
    console.log('  - localhost.key');
    process.exit(0);
}

console.log('Generating self-signed SSL certificates...');

try {
    // Install office-addin-dev-certs if not already installed
    console.log('Installing office-addin-dev-certs...');
    execSync('node node_modules/office-addin-dev-certs/cli.js install --days 365', {
        stdio: 'inherit',
        cwd: __dirname
    });
    
    // Copy certificates from user directory to project root
    const userCertDir = path.join(os.homedir(), '.office-addin-dev-certs');
    const sourceCert = path.join(userCertDir, 'localhost.crt');
    const sourceKey = path.join(userCertDir, 'localhost.key');
    
    if (fs.existsSync(sourceCert) && fs.existsSync(sourceKey)) {
        fs.copyFileSync(sourceCert, certPath);
        fs.copyFileSync(sourceKey, keyPath);
        console.log('\nCertificates generated and copied successfully!');
        console.log('  - localhost.crt');
        console.log('  - localhost.key');
    } else {
        console.error('\nWarning: Could not find certificates in user directory.');
        console.error('Please check:', userCertDir);
    }
} catch (error) {
    console.error('\nError generating certificates:', error.message);
    console.error('\nAlternatively, you can manually install office-addin-dev-certs globally:');
    console.error('  npm install -g office-addin-dev-certs');
    console.error('  office-addin-dev-certs install --days 365');
    console.error('Then copy the certificates to this folder:');
    console.error('  - localhost.crt');
    console.error('  - localhost.key');
    process.exit(1);
}
