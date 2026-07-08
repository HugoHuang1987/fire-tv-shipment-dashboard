# Copy this file to a local, untracked script and fill values on each machine.

$env:FIRE_TV_SHARED_ROOT = "D:\path\to\Fire TV shipment"
$env:FIRE_TV_RDM_BASE_URL = "https://your-rdm-host/rdm/sw/odf/listing.html"
$env:FIRE_TV_WIKI_URL = "https://your-wiki-host/wiki/pages/viewpage.action?pageId=..."

# Optional runtime paths. If omitted, scripts use python/node from PATH.
$env:FIRE_TV_PYTHON = "python"
$env:FIRE_TV_NODE = "node"
$env:FIRE_TV_NODE_MODULES = ""

# Optional notifications. Keep real values out of git.
$env:FIRE_TV_SERVERCHAN_SENDKEY = ""
$env:FIRE_TV_WECOM_WEBHOOK = ""

# Optional LLM connection. Keep real values out of git.
$env:FIRE_TV_LLM_BASE_URL = ""
$env:FIRE_TV_LLM_MODEL = ""
$env:FIRE_TV_LLM_API_KEY = ""

