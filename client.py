import os
import pathlib
import requests
from flask import Flask, render_template, request, session, abort, redirect
from google.oauth2 import id_token
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from pip._vendor import cachecontrol
import google.auth.transport.requests
import json


from sheetsBot import create

app = Flask(__name__)
app.secret_key = (os.getenv('SECRET_KEY')) # should match with what's in client_secret.json
uri = (os.getenv('URI'))

os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

GOOGLE_CLIENT_ID = (os.getenv('GOOGLE_CLIENT_ID'))
client_secrets_file = os.path.join(pathlib.Path(__file__).parent, "client_secret.json")
CALLBACK_URI = (os.getenv('CALLBACK_URI'))
SCOPES = ["openid", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/calendar.readonly", "https://www.googleapis.com/auth/userinfo.profile"]

flow = Flow.from_client_secrets_file(
    client_secrets_file=client_secrets_file, 
    scopes=SCOPES,
    redirect_uri=CALLBACK_URI
    )

def login_is_required(function):
    def wrapper(*args, **kwargs):
        if "google_id" not in session:
            return abort(401) # Auth required (client-side session package, stored in browser cookies. Don't use in production!)
        else:
            return function()
        
    # Renaming the function name:
    wrapper.__name__ = function.__name__
        
    return wrapper

@app.route("/")
def home():
    return render_template("login.html")

@app.route("/login")
def login():
    authorization_url, state = flow.authorization_url(access_type='offline', include_granted_scopes='true') # access_type='offline' will allow for a refresh token, which is needed for line 104
    session["state"] = state
    return redirect(authorization_url)

@app.route("/callback")
def callback():
    flow.fetch_token(authorization_response=request.url)

    if not session["state"] == request.args["state"]:
        abort(500) # State doesn't match

    credentials = flow.credentials
    request_session = requests.session()
    cached_session = cachecontrol.CacheControl(request_session)
    token_request = google.auth.transport.requests.Request(session=cached_session)

    id_info = id_token.verify_oauth2_token(
        id_token=credentials._id_token,
        request=token_request,
        audience=GOOGLE_CLIENT_ID
    )

    session["google_id"] = id_info.get("sub")
    session["full_name"] = id_info.get("name")
    session["first_name"] = id_info.get("given_name")
    session["credentials"] = credentials.to_json() # serialize & store

    return redirect("/protected_area")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/protected_area")
@login_is_required
def protected_area():
    first_name = session["first_name"]
    context = {
        'first_name': first_name,
        'uri': uri
    }

    return render_template("index.html", **context)

@app.route("/protected_area2", methods=["POST"])
@login_is_required
def protected_area2():
    if request.method == "POST":
        full_name = session["full_name"]
        pos = request.form["position"]
        creds = json.loads(session["credentials"])
        credentials = Credentials.from_authorized_user_info(creds, SCOPES)

        create(full_name, pos, credentials)
        return render_template('logout.html')


if __name__ == "__main__":
    app.run(debug=True)