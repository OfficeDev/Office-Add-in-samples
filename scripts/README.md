# Build Test Scripts

This directory contains scripts for testing Office Add-in sample builds.

## Files

### `test-builds.sh`

Main test script that builds all samples with package.json files.

**Features:**
- Tests all 69+ samples (including nested manifest-configurations)
- Tracks expected vs unexpected failures
- Generates JSON and Markdown reports
- Creates detailed logs for each sample
- Handles different build script variations (build, build:dev)

**Usage:**

```bash
# Run locally
./scripts/test-builds.sh

# Run with specific Node version
nvm use 20
./scripts/test-builds.sh
```

**Output files:**
- `build-test-results.json` - Detailed results in JSON format
- `build-test-summary.md` - Human-readable summary
- `build-logs/` - Individual log files for each sample

### `build-test-config.json`

Configuration file that defines:

- **skip**: Samples that should not be tested (e.g., .NET Blazor samples)
- **expectedFailures**: Samples with known build issues that are tracked but don't fail the workflow

**Example:**

```json
{
  "skip": [
    {
      "path": "Samples/blazor-add-in",
      "reason": ".NET Blazor samples require different build tooling",
      "pattern": "blazor-add-in"
    }
  ],
  "expectedFailures": [
    {
      "path": "Samples/auth/Office-Add-in-SSO-NAA/manifest-configurations/unified",
      "reason": "Webpack config references ./src/taskpane files that don't exist in nested config directory",
      "issue": "Pre-existing configuration issue"
    }
  ]
}
```

## GitHub Actions Workflow

The build tests run automatically on:
- **Pull requests to main** - Tests all samples
- **Manual trigger** - Via workflow_dispatch

### Workflow Features

1. **Matrix testing**: Tests against Node.js 20, 22, and 24
2. **Smart caching**: Caches npm downloads with restore-keys for partial hits
3. **Artifact upload**: Saves test results and logs for 30 days
4. **PR comments**: Posts results summary to pull requests
5. **Expected failure handling**: Only fails on unexpected failures

### Viewing Results

**In GitHub Actions:**
1. Go to Actions tab
2. Select "Build Tests" workflow
3. View run results and download artifacts

**In Pull Requests:**
- Bot automatically comments with test summary
- Click "Details" on the check to see full logs

## Adding New Expected Failures

If a sample has a known build issue:

1. Edit `scripts/build-test-config.json`
2. Add entry to `expectedFailures` array:
   ```json
   {
     "path": "Samples/your-sample-path",
     "reason": "Brief description of why it fails",
     "issue": "GitHub issue link or more details"
   }
   ```
3. Commit the change

The sample will still be tested, but won't fail the workflow.

## Skipping Samples

To skip a sample entirely (e.g., different tech stack):

1. Edit `scripts/build-test-config.json`
2. Add entry to `skip` array:
   ```json
   {
     "path": "Samples/sample-to-skip",
     "reason": "Why it should be skipped",
     "pattern": "pattern-to-match"
   }
   ```

## Troubleshooting

### Prerequisites

The test script requires:
- **Node.js** (version 20, 22, or 24)
- **jq** - JSON processor for parsing configuration files
  - Ubuntu/Debian: `sudo apt-get install jq`
  - macOS: `brew install jq`
  - Windows: Download from https://jqlang.github.io/jq/download/
  - GitHub Actions: Pre-installed on ubuntu-latest runners ✅

### Local testing

```bash
# Check prerequisites
node --version
npm --version
jq --version

# Run full test suite
./scripts/test-builds.sh

# Test a specific sample manually
cd Samples/excel-data-types-explorer
npm ci
npm run build

# View logs from last run
cat build-logs/Samples_excel-data-types-explorer.log
```

### CI/CD issues

**Cache problems:**
- Workflow uses restore-keys for partial cache hits
- Cache key includes Node version and all package-lock.json hashes
- Manual cache clearing: Re-run workflow with "Re-run all jobs"

**Timeout issues:**
- Per-sample timeout: 5 minutes (in script)
- Workflow timeout: 60 minutes (in workflow file)
- Adjust in `.github/workflows/build-tests.yml` if needed

**Node version issues:**
- Matrix tests Node 20, 22, 24
- Remove versions from matrix if not needed
- Ensure sample package.json has compatible engine requirements

## Maintenance

### Regular tasks

- **Review expected failures monthly** - Check if issues are resolved
- **Update Node versions** - Add new LTS versions to matrix
- **Monitor build times** - Optimize slow samples if needed

### After adding new samples

New samples are automatically detected and tested. No configuration needed unless:
- Sample uses different tech stack → Add to `skip`
- Sample has known build issue → Add to `expectedFailures`

## Performance

**Typical run times:**
- Full test suite (69 samples): ~15-25 minutes
- With cache hit: ~10-15 minutes
- Single sample: ~15-30 seconds

**Optimization tips:**
- npm cache reuses downloads across samples
- Parallel matrix testing (3 jobs) runs simultaneously
- Logs only kept for failed builds (configurable)
