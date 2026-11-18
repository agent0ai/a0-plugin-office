"""Register Office tokens on A0's existing authenticated VNC transport."""
from pathlib import Path
from urllib.parse import urlsplit
from starlette.responses import FileResponse, JSONResponse, PlainTextResponse
from starlette.websockets import WebSocket
from helpers import plugins, virtual_desktop
from plugins._browser.helpers.viewer_gateway import BrowserVncGateway

ASSETS = Path(__file__).resolve().parents[1] / 'webui'

class OfficeVncGateway(BrowserVncGateway):
    async def __call__(self, scope, receive, send):
        if scope['type'] == 'websocket':
            endpoint, path = self.endpoint(scope)
            if path != '/ws':
                await WebSocket(scope,receive=receive,send=send).close(code=1008)
                return
            await self.websocket(scope,receive,send)
        elif scope['type'] == 'http':
            await self.http(scope,receive,send)

    def allowed(self, scope):
        if plugins.get_toggle_state('office') != 'enabled' or not self.is_authorized(scope):
            return False
        headers = dict(scope.get('headers', []))
        if headers.get(b'sec-fetch-site', b'').lower() in {b'cross-site', b'same-site'}:
            return False
        origin = headers.get(b'origin', b'').decode('latin-1')
        if origin:
            try:
                parsed = urlsplit(origin)
            except ValueError:
                return False
            if parsed.scheme not in {'http', 'https'} or parsed.netloc.lower() != headers.get(b'host', b'').decode('latin-1').lower():
                return False
        elif scope['type'] == 'websocket':
            return False
        return True

    def endpoint(self, scope):
        route = self.session_request(self.relative_path(scope))
        if not route:
            return None, ''
        endpoint = virtual_desktop.proxy_for_token(route[0])
        if not endpoint or endpoint.owner != 'office-vnc' or endpoint.host != '127.0.0.1':
            return None, ''
        return endpoint, route[1]

    async def http(self, scope, receive, send):
        if not self.allowed(scope):
            await PlainTextResponse('Access denied',status_code=403)(scope,receive,send)
            return
        endpoint, path = self.endpoint(scope)
        headers = {'Cache-Control':'no-store','X-Frame-Options':'SAMEORIGIN','X-Content-Type-Options':'nosniff'}
        if not endpoint or scope.get('method') not in {'GET','HEAD'}:
            response = PlainTextResponse('Office viewer not found',status_code=404)
        elif path == '/credentials':
            response = JSONResponse({'password':endpoint.password},headers=headers)
        elif path in {'/index.html','/vnc.js'}:
            response = FileResponse(ASSETS / ('vnc.html' if path == '/index.html' else 'vnc.js'),headers=headers)
        else:
            response = PlainTextResponse('Not found',status_code=404)
        await response(scope,receive,send)


def register():
    # A0 v3 has a fixed VNC mount; dispatch only our tokens, including after a Hub install.
    if getattr(BrowserVncGateway, '_office_router', None):
        return
    original = BrowserVncGateway.__call__
    async def route(self, scope, receive, send):
        path = self.relative_path(scope)
        if path.startswith('/session/office-'):
            await OfficeVncGateway(self.flask_app)(scope,receive,send)
        else:
            await original(self,scope,receive,send)
    BrowserVncGateway._office_router = route
    BrowserVncGateway._office_original = original
    BrowserVncGateway.__call__ = route


def unregister():
    router = getattr(BrowserVncGateway, '_office_router', None)
    if router and BrowserVncGateway.__call__ is router:
        BrowserVncGateway.__call__ = BrowserVncGateway._office_original
        del BrowserVncGateway._office_router
        del BrowserVncGateway._office_original
