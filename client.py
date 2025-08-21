import os
import pathlib
import requests
from flask import Flask, render_template, request, session, abort, redirect
from google.oauth2 import id_token
from google_auth_oauthlib.flow import Flow
from pip._vendor import cachecontrol
import google.auth.transport.requests


from sheetsBot import create

app = Flask(__name__)
app.secret_key = (os.getenv('SECRET_KEY')) # should match with what's in client_secret.json

os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

GOOGLE_CLIENT_ID = (os.getenv('GOOGLE_CLIENT_ID'))
client_secrets_file = os.path.join(pathlib.Path(__file__).parent, "client_secret.json")
CALLBACK_URI = (os.getenv('CALLBACK_URI'))

flow = Flow.from_client_secrets_file(
    client_secrets_file=client_secrets_file, 
    scopes=["openid", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/calendar.readonly", "https://www.googleapis.com/auth/userinfo.profile"],
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
    authorization_url, state = flow.authorization_url()
    session["state"] = state
    return redirect(authorization_url)

@app.route("/callback")
def callback():
    print("req url: ", request.url)
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
    print("id_info: ", id_info)
    session["google_id"] = id_info.get("sub")
    session["full_name"] = id_info.get("name")
    session["first_name"] = id_info.get("given_name")
    return redirect("/protected_area")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/protected_area")
@login_is_required
def protected_area():
    full_name = session["full_name"]
    first_name = session["first_name"]
    context = {
        'full_name': full_name,
        'first_name': first_name
    }

    return render_template("index.html", **context)

@app.route("/protected_area2", methods=["POST"])
@login_is_required
def protected_area2():
    #return f"Open Google Sheets to view your timesheet. <a href='/logout'><button>Logout</button></a>"
    if request.method == "POST":
        full_name = session["full_name"]
        pos = request.form["position"]
        #print("pos: ", request.form["position"])
        create(full_name, pos)
        return render_template('logout.html')
    # return f"Your timesheet has been created. Open your Google Sheets to verify :) <br/> <a href='/logout'><button>Logout</button></a>"


if __name__ == "__main__":
    app.run(debug=True)