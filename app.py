"""Streamlit Cloud entrypoint.

The dashboard implementation lives in helpdesk.py. Keeping this wrapper lets
Streamlit Cloud use the conventional app.py entrypoint while preserving the
existing helpdesk.py filename.
"""

import helpdesk  # noqa: F401
