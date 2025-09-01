
import streamlit as st, os, json, base64

st.set_page_config(page_title="Credential Diagnostic", layout="centered")

st.title("Credential Diagnostic")

def show(obj, label):
    st.code(f"{label}:\n" + str(obj))

# Check secrets presence
has_sa = "gcp_service_account" in st.secrets or "service_account" in st.secrets
st.write("Has secrets:", has_sa)

if "gcp_service_account" in st.secrets:
    st.success("Found [gcp_service_account] in secrets")
    st.write("client_email:", st.secrets["gcp_service_account"].get("client_email"))
elif "service_account" in st.secrets:
    st.success("Found [service_account] in secrets")
    st.write("client_email:", st.secrets["service_account"].get("client_email"))
else:
    st.warning("No service account in secrets")

# Check ENV variables
env_raw = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON") or os.environ.get("SERVICE_ACCOUNT_JSON")
st.write("Has ENV JSON:", bool(env_raw))

# Check files
exists = []
for p in ("./service_account.json", "/mount/data/service_account.json", "/mnt/data/service_account.json"):
    if os.path.exists(p):
        exists.append(p)
st.write("Existing SA files:", exists or "None")

st.info("If at least one of the three checks above is True (secrets/env/file), the app should NOT ask to upload.")
