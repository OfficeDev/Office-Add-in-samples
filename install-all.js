const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function findPackageJsonDirs(dir, dirList = []) {
  const files = fs.readdirSync(dir);
  let hasPackageJson = false;

  files.forEach(file => {
    const filePath = path.join(dir, file);
    const stat = fs.statSync(filePath);

    if (file === 'package.json') {
      hasPackageJson = true;
    } else if (stat.isDirectory() && file !== 'node_modules' && file !== '.git') {
      findPackageJsonDirs(filePath, dirList);
    }
  });

  if (hasPackageJson) {
    dirList.push(dir);
  }

  return dirList;
}

function installDependencies(dir) {
  try {
    console.log(`\nInstalling: ${path.relative(process.cwd(), dir)}`);

    // Remove node_modules first for clean install
    const nodeModulesPath = path.join(dir, 'node_modules');
    if (fs.existsSync(nodeModulesPath)) {
      console.log('  Removing old node_modules...');
      execSync(`rmdir /s /q "${nodeModulesPath}"`, { stdio: 'ignore', shell: true });
    }

    // Run npm install
    execSync('npm install', {
      cwd: dir,
      stdio: 'pipe',
      timeout: 300000 // 5 minute timeout per install
    });

    console.log('  ✓ Installed successfully');
    return { success: true, dir };
  } catch (error) {
    console.log(`  ✗ Failed: ${error.message}`);
    return { success: false, dir, error: error.message };
  }
}

// Main execution
console.log('Finding directories with package.json...');
const dirs = findPackageJsonDirs('./Samples');
console.log(`Found ${dirs.length} directories to install\n`);

const results = {
  success: [],
  failed: []
};

dirs.forEach((dir, index) => {
  console.log(`[${index + 1}/${dirs.length}]`);
  const result = installDependencies(dir);

  if (result.success) {
    results.success.push(result.dir);
  } else {
    results.failed.push({ dir: result.dir, error: result.error });
  }
});

console.log(`\n=== Summary ===`);
console.log(`Successful: ${results.success.length}`);
console.log(`Failed: ${results.failed.length}`);

if (results.failed.length > 0) {
  console.log('\nFailed installations:');
  results.failed.forEach(f => {
    console.log(`  - ${path.relative(process.cwd(), f.dir)}: ${f.error}`);
  });
}
