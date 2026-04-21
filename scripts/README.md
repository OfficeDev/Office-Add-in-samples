# Build Test Suite

This suite provides automated build testing for all Office Add-in samples. Tests run on every PR to main to catch build regressions before they're merged.

---

## Quick Start

### For PR authors

**Your PR is automatically tested if it changes any of the following files:**

- TypeScript or JavaScript files in `Samples/**`
- `package.json` or `package-lock.json` files
- Webpack or TypeScript configurations
- The build test workflow itself

**What happens:** When your PR triggers testing, the following occurs:

1. Tests run against Node.js 20, 22, and 24 (in parallel).
1. A bot posts results to your PR (~15-20 minutes).
1. ✅ Green check = All builds passed or failed as expected.
1. ❌ Red X = An unexpected build failure occurred and needs fixing.

**If tests fail:** Follow these steps to diagnose and fix the issue:

1. Check the bot comment for which samples failed.
1. Download artifacts for detailed error logs.
1. Test locally: `cd Samples/<sample> && npm ci && npm run build`
1. Fix the issues and push your changes.

### For Maintainers

**New samples:** The system automatically detects and tests new samples, so you don't need to configure anything.

**If a sample has known build problems:** Add the sample to the `expectedFailures` array in `build-test-config.json`.

**If a sample should be skipped:** Add the sample to the `skip` array in `build-test-config.json`.

---

## What Gets Tested

### ✅ Included

The testing process includes the following types of samples:

- All root samples (for example, `Samples/excel-data-types-explorer/`)
- All nested manifest configurations
- Samples with `build` or `build:dev` npm scripts
- Both build scripts if they exist and are different

### ⊘ Excluded

The following items are excluded from testing:

- .NET Blazor samples (require .NET SDK, not Node.js)
- Samples without `package.json`
- Samples without build scripts
- `node_modules` directories

---

## Expected failures

These samples have **known build problems**. The workflow tracks these problems but doesn't fail because of them:

1. **Office-Add-in-SSO-NAA/manifest-configurations/unified**
   - Webpack config references `./src/taskpane` files that don't exist in nested directory

1. **Excel.OfflineStorageAddin/manifest-configurations/add-in-only**
   - Same webpack path problem

1. **word-citation-management/manifest-configurations/unified**
   - Same webpack path problem

The workflow still tests these problems so you know immediately if they get fixed.

---

## Understanding results

### ✅ Pass criteria

A sample passes when the following conditions are met:

- `npm ci` succeeds (clean dependency install).
- `npm run build:dev` succeeds (if script exists).
- `npm run build` succeeds (if script exists and different).

### ❌ Failure types

There are two types of build failures:

**Expected failure:**

- The sample is in the known failures list.
- The sample is still tested so you can track whether it's fixed.
- **This type of failure doesn't fail the workflow.**

**Unexpected failure:**

- The sample built successfully before but now fails.
- **This type of failure FAILS the workflow.**
- It requires investigation and fixing.

### 📊 PR check status

- **✅ Green**: All builds passed or failed as expected.
- **❌ Red**: One or more unexpected failures.
- **💬 Comment**: Bot posts detailed summary.

---

## Viewing results

### In PR comments

The bot automatically posts the following information:

- Total samples tested.
- Pass/fail counts.
- List of unexpected failures (if any).
- Collapsible list of passed samples.

### In GitHub Actions

To view detailed results in GitHub Actions, follow these steps:

1. Select **Details** next to the check.
1. View console output for each Node.js version.
1. Download artifacts (30-day retention):
   - `build-test-results.json` - Machine-readable results.
   - `build-test-summary.md` - Human-readable summary.
   - `build-logs/` - Individual logs for each sample.

### Example results

The following example shows the format of the JSON results file:

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

### Files in this directory

**`test-builds.sh`** - This file is the main test script. It performs the following actions:

- Tests all 69+ samples.
- Tracks expected versus unexpected failures.
- Generates JSON and Markdown reports.
- Creates detailed logs for each sample.

**`build-test-config.json`** - This file contains the configuration settings:

- `skip`: Samples to exclude from testing.
- `expectedFailures`: Known problems that don't fail the workflow.

### Configuration examples

The following examples show how to configure the test suite for different scenarios.

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

## Adding new samples

**Most samples (90%):** No action is needed because the process automatically detects and tests them.

**Sample has known build issue:** If your sample has a known build issue, follow these steps:

1. Edit `scripts/build-test-config.json`.
1. Add the sample to the `expectedFailures` array.
1. Commit your changes with your PR.

**Sample uses different tech stack:** If your sample uses a different technology stack (such as .NET), follow these steps:

1. Edit `scripts/build-test-config.json`.
1. Add the sample to the `skip` array.
1. Commit your changes with your PR.

For detailed examples, see the main repository.

---

## Local testing

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

### Common build failures

**TypeScript errors:**

- Fix type errors in your code.
- Check `tsconfig.json` configuration.

**Missing dependencies:**

- Verify `package.json` is correct.
- Run `npm ci` to ensure clean install.

**Webpack errors:**

- Check `webpack.config.js`.
- Verify entry points exist.

**Import errors:**

- Check file paths are correct.
- Ensure imports match file structure.

### CI/CD problems

**Cache problems:**

- Workflow uses restore keys for partial cache hits.
- Cache key includes Node version and all `package-lock.json` hashes.
- Clear cache: Re-run workflow by using **Re-run all jobs**.

**Timeout problems:**

- Workflow timeout: 60 minutes.
- Adjust in `.github/workflows/build-tests.yml` if needed.

**Node version problems:**

- Matrix tests Node 20, 22, 24.
- Ensure `package.json` has compatible engine requirements.

---

## Performance

**Typical run times:**

- Per PR: About 15-20 minutes (three Node versions in parallel)
- Per Node version: About 10-15 minutes
- With cache hit: About 8-12 minutes
- Single sample: About 15-30 seconds

**Cost:**

- Free for public repos ✅
- About 40-60 minutes of GitHub Actions time per PR

**Optimization:**

- npm cache reuses downloads across samples
- Parallel matrix testing (three jobs run simultaneously)
- Individual logs capture install/build output for all samples

---

## Maintenance

### Regular Tasks

- **Monthly**: Review expected failures - remove if fixed
- **Quarterly**: Update Node versions in matrix
- **As needed**: Optimize slow samples

### Workflow features

- **Matrix testing**: Node.js 20, 22, 24
- **Smart caching**: npm downloads with restore keys
- **Artifact upload**: Results and logs with 30-day retention
- **PR comments**: Auto-posted summaries
- **Expected failures**: Tracked separately from real failures

---

## FAQ

**Q: Can I run tests locally?**
A: Yes! Use `./scripts/test-builds.sh` (requires Node.js and jq).

**Q: Why test multiple Node versions?**
A: Testing multiple versions helps catch compatibility problems. Different versions might handle builds differently.

**Q: What if my sample doesn't have a build script?**
A: The process automatically skips it. Add a build script if you need one.

**Q: Can I skip my sample temporarily?**
A: Yes, add it to the `skip` section in `scripts/build-test-config.json` with a reason.

**Q: Tests are too slow in my PR**
A: Tests only run on file changes in `Samples/**`. Unrelated changes don't trigger tests.

**Q: How do I re-run failed tests?**
A: Click **Re-run failed jobs** in GitHub Actions, or push a new commit.

**Q: What if an expected failure now passes?**
A: 🎉 Great! Remove it from `expectedFailures` in the config.

---

## Support

- **Test suite problems**: Open an issue with the `ci/cd` label.
- **Sample-specific problems**: Open an issue that references the sample name.
- **Questions**: Ask in PR comments or check this README.
