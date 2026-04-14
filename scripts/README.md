# Build Test Suite

Automated build testing for all 69+ Office Add-in samples. Tests run on every PR to main, catching build regressions before merge.

---

## Quick Start

### For PR Authors

**Your PR will be automatically tested if it changes:**
- TypeScript or JavaScript files in `Samples/**`
- package.json or package-lock.json files
- Webpack or TypeScript configurations
- The build test workflow itself

**What happens:**
1. Tests run against Node.js 20, 22, and 24 (parallel)
2. Bot posts results to your PR (~15-20 minutes)
3. ✅ Green check = All builds passed or failed as expected
4. ❌ Red X = Unexpected build failure - needs fixing

**If tests fail:**
1. Check bot comment for which samples failed
2. Download artifacts for detailed error logs
3. Test locally: `cd Samples/<sample> && npm ci && npm run build`
4. Fix issues and push

### For Maintainers

**New samples:** Automatically detected and tested - no configuration needed!

**Sample has known build issues:** Add to `build-test-config.json` → `expectedFailures`

**Sample should be skipped:** Add to `build-test-config.json` → `skip`

---

## What Gets Tested

### ✅ Included
- All root samples (e.g., `Samples/excel-data-types-explorer/`)
- All nested manifest-configurations
- Samples with `build` or `build:dev` npm scripts
- Both build scripts if they exist and are different

### ⊘ Excluded
- .NET Blazor samples (require dotnet SDK, not Node.js)
- Samples without package.json
- Samples without build scripts
- node_modules directories

---

## Expected Failures

These samples have **known build issues** and are tracked but don't fail the workflow:

1. **Office-Add-in-SSO-NAA/manifest-configurations/unified**
   - Webpack config references `./src/taskpane` files that don't exist in nested directory

2. **Excel.OfflineStorageAddin/manifest-configurations/add-in-only**
   - Same webpack path issue

3. **word-citation-management/manifest-configurations/unified**
   - Same webpack path issue

These are still tested so we know immediately if they get fixed.

---

## Understanding Results

### ✅ Pass Criteria
- `npm ci` succeeds (clean dependency install)
- `npm run build:dev` succeeds (if script exists)
- `npm run build` succeeds (if script exists and different)

### ❌ Failure Types

**Expected failure:**
- Sample is in known failures list
- Still tested to track if fixed
- **Does NOT fail workflow**

**Unexpected failure:**
- Sample built successfully before but now fails
- **FAILS workflow**
- Requires investigation and fixing

### 📊 PR Check Status

- **✅ Green**: All builds passed or failed as expected
- **❌ Red**: One or more unexpected failures
- **💬 Comment**: Bot posts detailed summary

---

## Viewing Results

### In PR Comments

Bot automatically posts:
- Total samples tested
- Pass/fail counts
- List of unexpected failures (if any)
- Collapsible list of passed samples

### In GitHub Actions

1. Click "Details" next to the check
2. View console output for each Node.js version
3. Download artifacts (30-day retention):
   - `build-test-results.json` - Machine-readable results
   - `build-test-summary.md` - Human-readable summary
   - `build-logs/` - Individual logs for each sample

### Example Results

**build-test-results.json:**
```json
{
  "timestamp": "2026-04-14T17:30:00Z",
  "node_version": "v24.0.0",
  "total": 69,
  "passed": 65,
  "failed": 3,
  "expected_failures": 3,
  "unexpected_failures": 0,
  "skipped": 1
}
```

---

## Configuration

### Files in This Directory

**`test-builds.sh`** - Main test script
- Tests all 69+ samples
- Tracks expected vs unexpected failures
- Generates JSON and Markdown reports
- Creates detailed logs for each sample

**`build-test-config.json`** - Configuration
- `skip`: Samples to exclude from testing
- `expectedFailures`: Known issues that don't fail workflow

### Configuration Examples

**Skip a sample:**
```json
{
  "skip": [
    {
      "path": "Samples/blazor-add-in",
      "reason": ".NET Blazor samples require dotnet SDK",
      "pattern": "blazor-add-in"
    }
  ]
}
```

**Add expected failure:**
```json
{
  "expectedFailures": [
    {
      "path": "Samples/your-sample-name",
      "reason": "Brief description of why it fails",
      "issue": "GitHub issue link or more details"
    }
  ]
}
```

---

## Adding New Samples

**Most samples (90%):** No action needed - automatically detected and tested!

**Sample has known build issue:**
1. Edit `scripts/build-test-config.json`
2. Add to `expectedFailures` array
3. Commit with your PR

**Sample uses different tech stack:**
1. Edit `scripts/build-test-config.json`
2. Add to `skip` array
3. Commit with your PR

See the main repository for detailed examples.

---

## Local Testing

### Prerequisites

- **Node.js** (version 20, 22, or 24)
- **jq** - JSON processor
  - Ubuntu/Debian: `sudo apt-get install jq`
  - macOS: `brew install jq`
  - Windows: Download from https://jqlang.github.io/jq/download/
  - GitHub Actions: Pre-installed ✅

### Run Tests

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

---

## Troubleshooting

### Common Build Failures

**TypeScript errors:**
- Fix type errors in your code
- Check tsconfig.json configuration

**Missing dependencies:**
- Verify package.json is correct
- Run `npm ci` to ensure clean install

**Webpack errors:**
- Check webpack.config.js
- Verify entry points exist

**Import errors:**
- Check file paths are correct
- Ensure imports match file structure

### CI/CD Issues

**Cache problems:**
- Workflow uses restore-keys for partial cache hits
- Cache key includes Node version and all package-lock.json hashes
- Clear cache: Re-run workflow with "Re-run all jobs"

**Timeout issues:**
- Per-sample timeout: 5 minutes (in script)
- Workflow timeout: 60 minutes
- Adjust in `.github/workflows/build-tests.yml` if needed

**Node version issues:**
- Matrix tests Node 20, 22, 24
- Ensure package.json has compatible engine requirements

---

## Performance

**Typical run times:**
- Per PR: ~15-20 minutes (3 Node versions in parallel)
- Per Node version: ~10-15 minutes
- With cache hit: ~8-12 minutes
- Single sample: ~15-30 seconds

**Cost:**
- Free for public repos ✅
- ~40-60 minutes of GitHub Actions time per PR

**Optimization:**
- npm cache reuses downloads across samples
- Parallel matrix testing (3 jobs run simultaneously)
- Individual logs only for failed builds

---

## Maintenance

### Regular Tasks

- **Monthly**: Review expected failures - remove if fixed
- **Quarterly**: Update Node versions in matrix
- **As needed**: Optimize slow samples

### Workflow Features

- **Matrix testing**: Node.js 20, 22, 24
- **Smart caching**: npm downloads with restore-keys
- **Artifact upload**: Results and logs (30-day retention)
- **PR comments**: Auto-posted summaries
- **Expected failures**: Tracked separately from real failures

---

## FAQ

**Q: Can I run tests locally?**
A: Yes! `./scripts/test-builds.sh` (requires Node.js and jq)

**Q: Why test multiple Node versions?**
A: Catches compatibility issues. Different versions may have different build behavior.

**Q: What if my sample doesn't have a build script?**
A: It will be skipped automatically. Add a build script if needed.

**Q: Can I skip my sample temporarily?**
A: Yes, add to `skip` in `scripts/build-test-config.json` with a reason.

**Q: Tests are too slow in my PR**
A: Tests only run on file changes in `Samples/**`. Unrelated changes won't trigger tests.

**Q: How do I re-run failed tests?**
A: Click "Re-run failed jobs" in GitHub Actions, or push a new commit.

**Q: What if an expected failure now passes?**
A: 🎉 Great! Remove it from `expectedFailures` in the config.

---

## Support

- **Test suite issues**: Open issue with `ci/cd` label
- **Sample-specific issues**: Open issue referencing sample name
- **Questions**: Ask in PR comments or check this README
