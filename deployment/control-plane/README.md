# Windows companion control plane

This host is the durable HTTPS relay between the MCP controller and outbound
Windows workers. It has no arbitrary-command endpoint. The only accepted
requests are typed device health, action submission/status, polling, and
acknowledgement routes.

## Entra resources

Create or select a single-tenant Entra application that exposes an API, usually
with an Application ID URI such as `api://<control-plane-client-id>`.

- Set the resource application's `api.requestedAccessTokenVersion` manifest
  property to `2`. Clients request `api://<control-plane-client-id>/.default`,
  while the two server-side audience fields contain the same application's
  bare client-ID GUID because that is the `aud` claim in a v2 access token.
- Expose admin-restricted delegated scope `WindowsCompanion.Control` for the
  signed-in MCP controller.
- Expose application role `WindowsCompanion.Control` for a daemon controller,
  if used, with allowed member type `Applications`.
- Expose application role `WindowsCompanion.Device` for each certificate-backed
  worker application, with allowed member type `Applications`; do not expose it
  as a delegated scope.
- Assign only the controller permission to the MCP app and only the device role
  to worker identities. Grant tenant admin consent after review.

`deployment/entra/New-WindowsControlPlaneEntraResources.ps1` provisions this
separate resource application, declares the selected permission on an existing
controller app, creates or updates one certificate-backed worker app, and
writes the IDs needed by the JSON configurations. It creates no secret and
does not grant any delegated or application consent unless the explicit
`-GrantAdminConsent` switch is present. Run it once per separately credentialed
worker; do not reuse a worker private key across laptops.

The service validates an exact tenant, issuer, audience, RSA algorithm, and
pinned public JWKS. Copy the current Entra public signing keys into the local
configuration and rotate them before the old key expires. The sample signing
key has no retained private key and cannot authenticate anything.

## Configure

1. Install `microsoft-agent[control-plane]`.
2. Copy `control-plane.sample.json` to a protected file ending in
   `.local.control-plane.json`.
3. Replace the tenant, bare API-app client-ID audience, public JWKS, controller
   permissions, device identities, client IDs/subjects, certificate
   thumbprints, and action allowlists. Keep every device disabled until
   enrollment is complete.
4. Validate without opening a listener:

   ```bash
   microsoft-windows-control-plane \
     --config deployment/control-plane/control-plane.local.control-plane.json \
     --validate-config
   ```

5. Set `database_path` to
   `/var/lib/microsoft-agent-control-plane/windows-companion.db` for the
   supplied Linux service unit.

The default listener is loopback-only. A non-loopback bind needs the explicit
`allow_non_loopback_bind` flag. The process intentionally disables proxy-header
rewriting so the mTLS header is trusted only when the direct peer IP is in
`trusted_proxy_client_hosts`.

For the supplied Nginx and systemd assets, keep these exact values:

```json
{
  "bind_host": "127.0.0.1",
  "port": 8443,
  "allow_non_loopback_bind": false,
  "database_path": "/var/lib/microsoft-agent-control-plane/windows-companion.db",
  "control_plane": {
    "trusted_proxy_mtls_header": "X-Client-Cert-Thumbprint",
    "trusted_proxy_client_hosts": ["127.0.0.1"],
    "require_device_mtls": true
  }
}
```

The API app client ID goes in both bare GUID audience fields. Worker clients
request `api://<api-client-id>/.default`. Add the worker application client ID
to that device's `device_client_ids`, and optionally add its service-principal
object ID to `device_principal_subjects`. Use the uppercase SHA-1 fingerprint
of the mTLS leaf certificate in `certificate_thumbprint`; that is the
fingerprint Nginx exports after validating the certificate chain.

## Hardened Linux service

The checked-in unit assumes systemd 249 or later, a dedicated non-login user,
and a virtual environment at `/opt/microsoft-agent-control-plane/venv`. An
administrator can stage it with commands equivalent to the following after
reviewing package provenance and pinning the deployed release:

```bash
sudo useradd --system --home-dir /var/lib/microsoft-agent-control-plane \
  --shell /usr/sbin/nologin microsoft-agent-control-plane
sudo install -d -o root -g root -m 0755 /opt/microsoft-agent-control-plane
sudo python3 -m venv /opt/microsoft-agent-control-plane/venv
sudo /opt/microsoft-agent-control-plane/venv/bin/pip install \
  'microsoft-agent[control-plane]==<reviewed-version>'

sudo install -d -o root -g microsoft-agent-control-plane -m 0750 \
  /etc/microsoft-agent
sudo install -o root -g microsoft-agent-control-plane -m 0640 \
  control-plane.local.control-plane.json \
  /etc/microsoft-agent/control-plane.local.control-plane.json
sudo install -o root -g microsoft-agent-control-plane -m 0640 \
  systemd/microsoft-windows-control-plane.env \
  /etc/microsoft-agent/control-plane.env
sudo install -o root -g root -m 0644 \
  systemd/microsoft-windows-control-plane.service \
  /etc/systemd/system/microsoft-windows-control-plane.service

sudo systemctl daemon-reload
sudo systemctl enable --now microsoft-windows-control-plane.service
```

The unit validates configuration before every start, grants no Linux
capability, makes only the state directory writable, restricts network traffic
to loopback, and allows binding only TCP 8443. If you intentionally change the
loopback port, change `SocketBindAllow` and the Nginx upstream in the same
review. Check the effective sandbox and logs with:

```bash
systemd-analyze security microsoft-windows-control-plane.service
systemctl status microsoft-windows-control-plane.service
journalctl -u microsoft-windows-control-plane.service
```

## TLS and mTLS boundary

Terminate trusted TLS at a reverse proxy connected to the service over
loopback or a private network. The proxy must verify the worker client
certificate, remove any client-supplied copy of the configured thumbprint
header, then inject the verified certificate thumbprint. Do not publish the
plain ASGI port directly.

The example at `nginx/windows-control-plane.conf` keeps client-certificate
verification optional at the TLS handshake so human controllers can connect,
then requires a successfully validated certificate on the fixed worker
`/relay/` routes. It replaces the trusted thumbprint header with Nginx's own
TLS-session fingerprint on every request, so a caller-supplied value cannot
reach the service. Replace all example hostnames and certificate paths, install
the server chain/key and companion issuing CA, then validate before reload:

```bash
sudo install -o root -g root -m 0644 \
  nginx/windows-control-plane.conf \
  /etc/nginx/conf.d/windows-control-plane.conf
sudo nginx -t
sudo systemctl reload nginx
```

Expose only Nginx TCP 443 through the host firewall. Keep TCP 8443 loopback
only. The worker certificate must chain to `companion-client-ca.pem`; use a CRL
or short certificate lifetime under your PKI policy, and rotate the Entra and
mTLS public credentials together when the same leaf certificate is used.

Back up the SQLite database and its `-wal`/`-shm` files as one consistent unit.
The service enforces record, queue, request, result, response, retention, and
database-size limits from the configuration.
