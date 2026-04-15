#!/bin/bash

# Build test script for Office Add-in samples
# Tests all samples with package.json files, tracking expected vs unexpected failures

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Counters
TOTAL=0
PASSED=0
FAILED=0
SKIPPED=0
EXPECTED_FAILURES=0
UNEXPECTED_FAILURES=0

# Arrays to track results
PASSED_SAMPLES=()
FAILED_SAMPLES=()
SKIPPED_SAMPLES=()
EXPECTED_FAILURE_SAMPLES=()
UNEXPECTED_FAILURE_SAMPLES=()

# Base directory
BASE_DIR=$(pwd)
SAMPLES_DIR="$BASE_DIR/Samples"
CONFIG_FILE="$BASE_DIR/scripts/build-test-config.json"
RESULTS_FILE="$BASE_DIR/build-test-results.json"
SUMMARY_FILE="$BASE_DIR/build-test-summary.md"
LOG_DIR="$BASE_DIR/build-logs"

# Create log directory
mkdir -p "$LOG_DIR"

echo -e "${BLUE}================================================${NC}"
echo -e "${BLUE}  Office Add-in Samples - Build Test Suite${NC}"
echo -e "${BLUE}================================================${NC}"
echo ""
echo "Base directory: $BASE_DIR"
echo "Node version: $(node --version)"
echo "npm version: $(npm --version)"
echo ""

# Load configuration
if [ ! -f "$CONFIG_FILE" ]; then
  echo -e "${RED}Error: Configuration file not found: $CONFIG_FILE${NC}"
  exit 1
fi

# Function to check if a path should be skipped
should_skip() {
  local sample_path=$1
  local skip_count=$(jq -r '.skip | length' "$CONFIG_FILE")

  for ((i=0; i<skip_count; i++)); do
    local pattern=$(jq -r ".skip[$i].pattern" "$CONFIG_FILE")
    if [[ "$sample_path" == *"$pattern"* ]]; then
      local reason=$(jq -r ".skip[$i].reason" "$CONFIG_FILE")
      echo "$reason"
      return 0
    fi
  done

  return 1
}

# Function to check if a failure is expected
is_expected_failure() {
  local sample_path=$1
  local expected_count=$(jq -r '.expectedFailures | length' "$CONFIG_FILE")

  for ((i=0; i<expected_count; i++)); do
    local pattern=$(jq -r ".expectedFailures[$i].pattern // .expectedFailures[$i].path" "$CONFIG_FILE")
    if [[ "$sample_path" == *"$pattern"* ]]; then
      return 0
    fi
  done

  return 1
}

# Function to get expected failure reason
get_expected_failure_reason() {
  local sample_path=$1
  local expected_count=$(jq -r '.expectedFailures | length' "$CONFIG_FILE")

  for ((i=0; i<expected_count; i++)); do
    local pattern=$(jq -r ".expectedFailures[$i].pattern // .expectedFailures[$i].path" "$CONFIG_FILE")
    if [[ "$sample_path" == *"$pattern"* ]]; then
      jq -r ".expectedFailures[$i].reason" "$CONFIG_FILE"
      return
    fi
  done
}

# Function to test a single sample
test_sample() {
  local sample_dir=$1
  local relative_path=${sample_dir#$BASE_DIR/}

  TOTAL=$((TOTAL + 1))
  echo ""
  echo -e "${BLUE}[$TOTAL] Testing: $relative_path${NC}"

  # Check if should skip
  if skip_reason=$(should_skip "$relative_path"); then
    echo -e "  ${YELLOW}⊘ SKIPPED${NC}: $skip_reason"
    SKIPPED=$((SKIPPED + 1))
    SKIPPED_SAMPLES+=("$relative_path")
    return
  fi

  # Check if package.json exists
  if [ ! -f "$sample_dir/package.json" ]; then
    echo -e "  ${YELLOW}⊘ SKIPPED${NC}: No package.json found"
    SKIPPED=$((SKIPPED + 1))
    SKIPPED_SAMPLES+=("$relative_path")
    return
  fi

  # Create log file for this sample
  local log_file="$LOG_DIR/$(echo "$relative_path" | tr '/' '_').log"

  cd "$sample_dir" || return

  # Install dependencies
  echo "  📦 Installing dependencies..."
  if ! npm ci --ignore-scripts > "$log_file" 2>&1; then
    echo -e "  ${RED}✗ FAILED${NC}: npm ci failed"
    FAILED=$((FAILED + 1))
    FAILED_SAMPLES+=("$relative_path (npm ci failed)")

    # npm ci failures are always unexpected (not build failures)
    # expectedFailures are for build-time webpack issues, not install issues
    UNEXPECTED_FAILURES=$((UNEXPECTED_FAILURES + 1))
    UNEXPECTED_FAILURE_SAMPLES+=("$relative_path")
    echo -e "  ${RED}  ⚠ UNEXPECTED FAILURE - install should not fail${NC}"

    cd "$BASE_DIR"
    return
  fi

  # Check for build scripts
  local has_build=$(jq -r '.scripts.build // empty' package.json)
  local has_build_dev=$(jq -r '.scripts."build:dev" // empty' package.json)

  if [ -z "$has_build" ] && [ -z "$has_build_dev" ]; then
    echo -e "  ${YELLOW}⊘ SKIPPED${NC}: No build scripts found"
    SKIPPED=$((SKIPPED + 1))
    SKIPPED_SAMPLES+=("$relative_path")
    cd "$BASE_DIR"
    return
  fi

  # Run builds
  local build_failed=false

  # Try build:dev first
  if [ -n "$has_build_dev" ]; then
    echo "  🔨 Running build:dev..."
    if ! npm run build:dev >> "$log_file" 2>&1; then
      echo -e "  ${RED}✗ FAILED${NC}: build:dev failed"
      build_failed=true
    else
      echo -e "  ${GREEN}✓${NC} build:dev passed"
    fi
  fi

  # Try build (if it exists and is different from build:dev)
  if [ -n "$has_build" ] && [ "$has_build" != "$has_build_dev" ]; then
    echo "  🔨 Running build..."
    if ! npm run build >> "$log_file" 2>&1; then
      echo -e "  ${RED}✗ FAILED${NC}: build failed"
      build_failed=true
    else
      echo -e "  ${GREEN}✓${NC} build passed"
    fi
  fi

  # Record results
  if [ "$build_failed" = true ]; then
    FAILED=$((FAILED + 1))
    FAILED_SAMPLES+=("$relative_path")

    if is_expected_failure "$relative_path"; then
      EXPECTED_FAILURES=$((EXPECTED_FAILURES + 1))
      EXPECTED_FAILURE_SAMPLES+=("$relative_path")
      reason=$(get_expected_failure_reason "$relative_path")
      echo -e "  ${YELLOW}✓ Expected failure confirmed${NC}"
      echo -e "  ${YELLOW}  Reason: $reason${NC}"
    else
      UNEXPECTED_FAILURES=$((UNEXPECTED_FAILURES + 1))
      UNEXPECTED_FAILURE_SAMPLES+=("$relative_path")
      echo -e "  ${RED}⚠ UNEXPECTED FAILURE - This should not fail!${NC}"
      echo -e "  ${RED}  See log: $log_file${NC}"
    fi
  else
    echo -e "  ${GREEN}✅ PASSED${NC}"
    PASSED=$((PASSED + 1))
    PASSED_SAMPLES+=("$relative_path")
  fi

  cd "$BASE_DIR"
}

# Find all samples with package.json
echo -e "${BLUE}Finding all samples with package.json...${NC}"
SAMPLE_DIRS=()
while IFS= read -r -d '' package_file; do
  sample_dir=$(dirname "$package_file")
  SAMPLE_DIRS+=("$sample_dir")
done < <(find "$SAMPLES_DIR" -name "node_modules" -prune -o -name "package.json" -type f -print0)

echo "Found ${#SAMPLE_DIRS[@]} samples to test"
echo ""

# Test each sample
for sample_dir in "${SAMPLE_DIRS[@]}"; do
  test_sample "$sample_dir"
done

# Generate results
echo ""
echo -e "${BLUE}================================================${NC}"
echo -e "${BLUE}  Test Results Summary${NC}"
echo -e "${BLUE}================================================${NC}"
echo ""
echo -e "Total samples tested:     ${BLUE}$TOTAL${NC}"
echo -e "✅ Passed:                ${GREEN}$PASSED${NC}"
echo -e "❌ Failed:                ${RED}$FAILED${NC}"
echo -e "  ├─ Expected failures:   ${YELLOW}$EXPECTED_FAILURES${NC}"
echo -e "  └─ Unexpected failures: ${RED}$UNEXPECTED_FAILURES${NC}"
echo -e "⊘  Skipped:               ${YELLOW}$SKIPPED${NC}"
echo ""

# Generate JSON results
cat > "$RESULTS_FILE" << EOF
{
  "timestamp": "$(date -u +%Y-%m-%dT%H:%M:%SZ)",
  "node_version": "$(node --version)",
  "total": $TOTAL,
  "passed": $PASSED,
  "failed": $FAILED,
  "expected_failures": $EXPECTED_FAILURES,
  "unexpected_failures": $UNEXPECTED_FAILURES,
  "skipped": $SKIPPED,
  "passed_samples": $(printf '%s\n' "${PASSED_SAMPLES[@]}" | jq -R . | jq -s .),
  "failed_samples": $(printf '%s\n' "${FAILED_SAMPLES[@]}" | jq -R . | jq -s .),
  "expected_failure_samples": $(printf '%s\n' "${EXPECTED_FAILURE_SAMPLES[@]}" | jq -R . | jq -s .),
  "unexpected_failure_samples": $(printf '%s\n' "${UNEXPECTED_FAILURE_SAMPLES[@]}" | jq -R . | jq -s .),
  "skipped_samples": $(printf '%s\n' "${SKIPPED_SAMPLES[@]}" | jq -R . | jq -s .)
}
EOF

# Generate Markdown summary
cat > "$SUMMARY_FILE" << EOF
### Summary

- **Total**: $TOTAL samples tested
- ✅ **Passed**: $PASSED
- ❌ **Failed**: $FAILED
  - Expected failures: $EXPECTED_FAILURES
  - ⚠️ **Unexpected failures**: $UNEXPECTED_FAILURES
- ⊘ **Skipped**: $SKIPPED

EOF

if [ $UNEXPECTED_FAILURES -gt 0 ]; then
  cat >> "$SUMMARY_FILE" << EOF
### ⚠️ Unexpected Failures

The following samples failed unexpectedly and need investigation:

EOF
  for sample in "${UNEXPECTED_FAILURE_SAMPLES[@]}"; do
    echo "- \`$sample\`" >> "$SUMMARY_FILE"
  done
  echo "" >> "$SUMMARY_FILE"
fi

if [ $EXPECTED_FAILURES -gt 0 ]; then
  cat >> "$SUMMARY_FILE" << EOF
### Expected Failures

The following samples failed as expected (known issues):

EOF
  for sample in "${EXPECTED_FAILURE_SAMPLES[@]}"; do
    reason=$(get_expected_failure_reason "$sample")
    echo "- \`$sample\` - $reason" >> "$SUMMARY_FILE"
  done
  echo "" >> "$SUMMARY_FILE"
fi

if [ $PASSED -gt 0 ]; then
  cat >> "$SUMMARY_FILE" << EOF
<details>
<summary>✅ Passed samples ($PASSED)</summary>

EOF
  for sample in "${PASSED_SAMPLES[@]}"; do
    echo "- \`$sample\`" >> "$SUMMARY_FILE"
  done
  echo "" >> "$SUMMARY_FILE"
  echo "</details>" >> "$SUMMARY_FILE"
  echo "" >> "$SUMMARY_FILE"
fi

cat >> "$SUMMARY_FILE" << EOF
---

📊 **Build logs**: Available in workflow artifacts

EOF

# Display summary
cat "$SUMMARY_FILE"

# Exit with appropriate code
if [ $UNEXPECTED_FAILURES -gt 0 ]; then
  echo -e "${RED}❌ Build tests failed: $UNEXPECTED_FAILURES unexpected failure(s)${NC}"
  exit 1
else
  echo -e "${GREEN}✅ All builds passed or failed as expected${NC}"
  exit 0
fi
