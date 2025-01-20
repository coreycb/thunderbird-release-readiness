# thunderbird-release-readiness
Script used to gather data for release readiness metrics.

## how to
In order to run this script you will need to export a bugzilla.mozilla.org
API key to the BMO_API_KEY environment variable.

### setup
```
python3 -m virtualenv .venv
source .venv/bin/activate
pip3 install -r requirements.txt
export PATH=$PATH:.
export BMO_API_KEY='<your API key>'
```

### execute
```
get-metrics
```
