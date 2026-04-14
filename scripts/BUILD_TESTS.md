# Build Tests - Quick Reference

## What This Tests

The build test suite automatically compiles and builds **all 69+ Office Add-in samples** to catch build regressions.

## When Tests Run

✅ **Automatically on every PR to main** that changes:
- TypeScript or JavaScript files
- package.json or package-lock.json files
- Webpack or TypeScript configurations
- The test workflow itself

✅ **Manual trigger** via GitHub Actions UI

## Test Matrix

Tests run against **3 Node.js versions** in parallel:
- Node.js 20 (LTS)
- Node.js 22 (LTS)
- Node.js 24 (Current)

## What Gets Tested

### Included ✅
- All root samples (e.g., `Samples/excel-data-types-explorer/`)
- All nested manifest-configurations (e.g., `Samples/auth/Office-Add-in-SSO-NAA/manifest-configurations/unified/`)
- Samples with either `build` or `build:dev` npm scripts
- Both build scripts if they exist and are different

### Excluded ⊘
- .NET Blazor samples (require dotnet SDK, not Node.js)
- Samples without package.json
- Samples without build scripts
- node_modules directories

## Expected Failures

Some samples have **known build issues** that are tracked but don't fail the workflow:

1. **Office-Add-in-SSO-NAA/manifest-configurations/unified**
   - Webpack config references source files that don't exist in nested directory
   - Pre-existing configuration issue

2. **Excel.OfflineStorageAddin/manifest-configurations/add-in-only**
   - Same webpack path issue as above

3. **word-citation-management/manifest-configurations/unified**
   - Same webpack path issue as above

These samples are still tested so we know if they get fixed.

## Understanding Test Results

### ✅ Pass Criteria
- `npm ci` succeeds (dependencies install cleanly)
- `npm run build:dev` succeeds (if script exists)
- `npm run build` succeeds (if script exists and different from build:dev)

### ❌ Failure Types

**Expected failure:**
- Sample is in the known failures list
- Still tested to track if it gets fixed
- **Does NOT fail the workflow**

**Unexpected failure:**
- Sample built successfully before but now fails
- **FAILS the workflow**
- Requires investigation and fixing

### 📊 PR Check Status

- **✅ Green check**: All samples passed or failed as expected
- **❌ Red X**: One or more unexpected failures
- **💬 Bot comment**: Detailed summary posted to PR

## Viewing Results

### In PR Comments
A bot automatically posts a summary with:
- Total samples tested
- Pass/fail counts
- List of unexpected failures (if any)
- Collapsible list of passed samples

### In GitHub Actions
1. Click "Details" next to the failed check
2. View console output for each Node.js version
3. Download artifacts for:
   - `build-test-results.json` - Machine-readable results
   - `build-test-summary.md` - Human-readable summary
   - `build-logs/` - Individual logs for each sample

### Example Artifact Contents

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

**build-logs/:**
- `Samples_excel-data-types-explorer.log`
- `Samples_word-citation-management.log`
- etc.

## Troubleshooting Failed Builds

### If your PR fails build tests:

1. **Check the bot comment** on your PR for which sample(s) failed

2. **Download the artifacts** to see detailed error logs

3. **Test locally:**
   ```bash
   cd Samples/<failed-sample>
   npm ci
   npm run build
   ```

4. **Common issues:**
   - **TypeScript errors**: Fix type errors in your changes
   - **Missing dependencies**: Update package.json
   - **Webpack errors**: Check webpack.config.js
   - **Import errors**: Verify file paths and imports

5. **Fix and push**: Once fixed locally, push your changes

### If an expected failure now passes:

🎉 **Great!** Remove it from `scripts/build-test-config.json` → `expectedFailures`

### If you need to add an expected failure:

See `scripts/README.md` for instructions on updating the config.

## Performance

**Typical run times:**
- **Per PR**: ~15-20 minutes (all 3 Node versions in parallel)
- **Per Node version**: ~10-15 minutes
- **With cache hit**: ~8-12 minutes

**Cost:**
- Free for public repos (Office-Add-in-samples is public)
- ~40-60 minutes of GitHub Actions time per PR

## Configuration Files

| File | Purpose |
|------|---------|
| `.github/workflows/build-tests.yml` | GitHub Actions workflow |
| `scripts/test-builds.sh` | Main test script |
| `scripts/build-test-config.json` | Expected failures & skip list |
| `scripts/README.md` | Detailed documentation |

## FAQ

**Q: Can I run tests locally?**
A: Yes! `./scripts/test-builds.sh` (requires Node.js and jq)

**Q: Why test multiple Node versions?**
A: Catches compatibility issues early. Different versions may have different build behavior.

**Q: What if my sample doesn't have a build script?**
A: It will be skipped automatically. Add a build script if needed.

**Q: Can I skip my sample temporarily?**
A: Yes, add it to `skip` in `scripts/build-test-config.json` with a reason.

**Q: Tests are too slow in my PR**
A: Tests only run on file changes in `Samples/**`. Change unrelated files to skip tests.

**Q: How do I re-run failed tests?**
A: Click "Re-run failed jobs" in GitHub Actions, or push a new commit.

## Support

- **Issues with the test suite itself**: Open an issue with the `ci/cd` label
- **Issues with a specific sample**: Open an issue referencing the sample name
- **Questions**: Check `scripts/README.md` or ask in PR comments
