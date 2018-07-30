#!/usr/local/bin/python3
# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
# See LICENSE in the project root for license information.
import base64
import mimetypes
import os
import pprint
import uuid
import json
import html
import urllib.parse

import flask
from flask_oauthlib.client import OAuth

import config

APP = flask.Flask(__name__, template_folder='static/templates')
APP.debug = True
APP.secret_key = 'development'
OAUTH = OAuth(APP)
MSGRAPH = OAUTH.remote_app(
    'microsoft',
    consumer_key=config.CLIENT_ID,
    consumer_secret=config.CLIENT_SECRET,
    request_token_params={'scope': config.SCOPES},
    base_url=config.RESOURCE + config.API_VERSION + '/',
    request_token_url=None,
    access_token_method='POST',
    access_token_url=config.AUTHORITY_URL + config.TOKEN_ENDPOINT,
    authorize_url=config.AUTHORITY_URL + config.AUTH_ENDPOINT)

@APP.route('/')
def homepage():
    return flask.render_template('homepage.html')

@APP.route('/login')
def login():
    flask.session['state'] = str(uuid.uuid4())
    return MSGRAPH.authorize(callback=config.REDIRECT_URI, state=flask.session['state'])

@APP.route('/login/authorized')
def authorized():
    """Handler for the application's Redirect Uri."""
    if str(flask.session['state']) != str(flask.request.args['state']):
        raise Exception('state returned to redirect URL does not match!')
    response = MSGRAPH.authorized_response()
    flask.session['access_token'] = response['access_token']
    return flask.redirect('/preexport')

@APP.route('/preexport')
def preexport():
    user_profile = MSGRAPH.get('me?$select=displayName,userPrincipalName', headers=request_headers()).data
    return flask.render_template('preexport.html',
                                 name=user_profile['displayName'],
                                 email=user_profile['userPrincipalName'])

def graph_generator(session, endpoint=None):
    """Generator for paginated result sets returned by Microsoft Graph.
    session = authenticated Graph session object
    endpoint = the Graph endpoint (for example, 'me/messages' for messages,
               or 'me/drive/root/children' for OneDrive drive items)
    """
    while endpoint:
        print('Retrieving next page ...')
        response = session.get(endpoint).json()
        yield from response.get('value')
        endpoint = response.get('@odata.nextLink')

def get_notebook(notebook_id, notebook_name):
    print(f"- {notebook_name}")
    return {
            'id': notebook_id,
            'name': notebook_name,
            'sections': [
                get_section(section['id'], section['displayName'])
                for section in MSGRAPH.get(f"me/onenote/notebooks/{notebook_id}/sections?$select=id,displayName").data['value']
                if section['displayName'] == 'Done'
                ]
            }

def get_section(section_id, section_name):
    print(f"- - {section_name}")
    return {
            'id': section_id,
            'name': section_name,
            'pages': [
                get_page(page['id'], page['title'])
                for page in MSGRAPH.get(f"me/onenote/sections/{section_id}/pages?$select=id,title").data['value']
                ]
            }

def get_page(page_id, page_title):
    print(f"- - - {page_title}")
    return {
            'id': page_id,
            'title': page_title,
            'content': MSGRAPH.get(f"me/onenote/pages/{page_id}/content").data.decode("utf-8", "strict")
            }

@APP.route('/export')
def export():
    return flask.render_template('exported.html', notebooks=[
        get_notebook(notebook['id'], notebook['displayName'])
        for notebook in MSGRAPH.get("me/onenote/notebooks?$select=id,displayName").data['value']
        if notebook['displayName'] == 'Archive'
        ]
        )

@MSGRAPH.tokengetter
def get_token():
    """Called by flask_oauthlib.client to retrieve current access token."""
    return (flask.session.get('access_token'), '')

def request_headers(headers=None):
    """Return dictionary of default HTTP headers for Graph API calls.
    Optional argument is other headers to merge/override defaults."""
    default_headers = {'SdkVersion': 'export-onenote',
                       'x-client-SKU': 'export-onenote',
                       'client-request-id': str(uuid.uuid4()),
                       'return-client-request-id': 'true'}
    if headers:
        default_headers.update(headers)
    return default_headers

if __name__ == '__main__':
    APP.run()
