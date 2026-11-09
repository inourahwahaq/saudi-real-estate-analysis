"""Download the six quarterly real estate files from the Saudi Open Data Platform.

usage: python3 fetch_realestate.py <raw_dir>
"""
import json
import sys
import urllib.parse
import urllib.request
from pathlib import Path

API = "https://open.data.gov.sa/data/api/datasets/resources?version=-1&dataset="
DATASETS = {
    "2025_Q1": "0fd9a088-8bd9-4d8a-8d69-63eac103238d",
    "2025_Q2": "ad218919-2014-4917-a85d-d4ec1a43c050",
    "2025_Q3": "5948497a-d84f-45a4-944c-50c59cff9629",
    "2025_Q4": "596f7335-afa5-4cb3-80a3-a1a67b947c35",
    "2026_Q1": "543ceeeb-4c3d-410e-a43a-dd2bdef1ab3c",
    "2026_Q2": "34e1da80-d4cd-4d65-bdc8-a8d48a7e7ee2",
}


def get(url):
    request = urllib.request.Request(url, headers={"User-Agent": "data-analysis"})
    with urllib.request.urlopen(request, timeout=120) as response:
        return response.read()


raw_dir = Path(sys.argv[1])
raw_dir.mkdir(parents=True, exist_ok=True)
for quarter, dataset in DATASETS.items():
    resources = json.loads(get(API + dataset))["resources"]
    url = next(r["downloadUrl"] for r in resources if r["format"] == "CSV")
    (raw_dir / f"{quarter}.csv").write_bytes(get(urllib.parse.quote(url, safe=":/")))
