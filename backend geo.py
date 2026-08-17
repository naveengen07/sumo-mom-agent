"""Compatibility launcher for local development.

The deployable Flask application lives in app.py. This file keeps the original
local command (`python "backend geo.py"`) working.
"""

from app import app


if __name__ == "__main__":
    import os

    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=os.getenv("FLASK_DEBUG") == "1")
