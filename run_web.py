"""Start the CMC Report Automation web server."""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from app.web.server import run

if __name__ == "__main__":
    run(host="0.0.0.0", port=5000, debug=False)
