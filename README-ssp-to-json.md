# SSP-to-JSON

This file describes `ssp-to-json.py`, which was adapted from `ssp-parse.py`.

The text analysis (matrix of difference measures betwee implementation narratives) has been removed, and the parsed controls from the SSP are dumped to stdout in JSON.

With the text analysis code removed, the only dependency is `python-docx`.  There is a requirements file you can use, `requirements-ssp-to-json.txt`.

## Installation

### Virtualenv Setup (optional)

```bash
virtualenv -p python3 venv
source venv/bin/activate
```

After you are finished running the script, you may deactivate the virtual environment if you wish.

```bash
deactivate
```

### Install Dependencies

```bash
pip3 install -r requirements-ssp-to-json.txt
```

## Running

Download an SSP (.docx format) based on the [FedRamp template](https://www.fedramp.gov/templates/).

Execute the script, with your choice of input and output filenames.

```bash
./ssp-to-json.py ssp.docx >ssp.json
```

