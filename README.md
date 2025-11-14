# thunderbird-release-readiness
Script used to gather data for release readiness metrics.

## pipx (recommended for isolated installs)

```bash
# Install from the local project directory
pipx install .

# Or install directly from GitHub (replace with your repo URL if different)
pipx install git+https://github.com/coreycb/thunderbird-release-readiness.git

# Run the tool
get-metrics --help

# When running the actual queries, set your Bugzilla API key
export BMO_API_KEY='<your API key>'
get-metrics
```

To upgrade later:

```bash
pipx upgrade thunderbird-release-readiness
```
