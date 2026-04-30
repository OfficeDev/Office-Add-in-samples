const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Security overrides to add
const SECURITY_OVERRIDES = {
  "uuid": "^14.0.0",
  "diff": "^9.0.0",
  "serialize-javascript": "^7.0.5",
  "tmp": "^0.2.5"
};

function findPackageJsonFiles(dir, fileList = []) {
  const files = fs.readdirSync(dir);

  files.forEach(file => {
    const filePath = path.join(dir, file);
    const stat = fs.statSync(filePath);

    if (stat.isDirectory() && file !== 'node_modules' && file !== '.git') {
      findPackageJsonFiles(filePath, fileList);
    } else if (file === 'package.json') {
      fileList.push(filePath);
    }
  });

  return fileList;
}

function updatePackageJson(filePath) {
  try {
    const content = fs.readFileSync(filePath, 'utf8');
    const pkg = JSON.parse(content);

    // Skip if no dependencies or devDependencies (likely a package inside node_modules)
    if (!pkg.dependencies && !pkg.devDependencies) {
      return { updated: false, reason: 'No dependencies' };
    }

    // Initialize overrides if it doesn't exist
    if (!pkg.overrides) {
      pkg.overrides = {};
    }

    // Add security overrides
    let changed = false;
    for (const [dep, version] of Object.entries(SECURITY_OVERRIDES)) {
      if (!pkg.overrides[dep] || pkg.overrides[dep] !== version) {
        pkg.overrides[dep] = version;
        changed = true;
      }
    }

    if (changed) {
      // Write back with proper formatting
      fs.writeFileSync(filePath, JSON.stringify(pkg, null, 2) + '\n', 'utf8');
      return { updated: true, reason: 'Overrides added' };
    }

    return { updated: false, reason: 'Already up to date' };
  } catch (error) {
    return { updated: false, reason: `Error: ${error.message}` };
  }
}

// Main execution
console.log('Finding package.json files in Samples directory...');
const packageFiles = findPackageJsonFiles('./Samples');
console.log(`Found ${packageFiles.length} package.json files\n`);

let updated = 0;
let skipped = 0;
let errors = 0;

packageFiles.forEach((file, index) => {
  const result = updatePackageJson(file);
  const relativePath = path.relative(process.cwd(), file);

  if (result.updated) {
    console.log(`✓ [${index + 1}/${packageFiles.length}] Updated: ${relativePath}`);
    updated++;
  } else {
    if (result.reason.startsWith('Error')) {
      console.log(`✗ [${index + 1}/${packageFiles.length}] ${result.reason}: ${relativePath}`);
      errors++;
    } else {
      // Uncomment to see skipped files
      // console.log(`- [${index + 1}/${packageFiles.length}] ${result.reason}: ${relativePath}`);
      skipped++;
    }
  }
});

console.log(`\n=== Summary ===`);
console.log(`Updated: ${updated}`);
console.log(`Skipped: ${skipped}`);
console.log(`Errors: ${errors}`);
console.log(`Total: ${packageFiles.length}`);
