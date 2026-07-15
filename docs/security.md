# Security contract

- Default input limit: 100 MiB.
- Default ZIP entry limit: 20,000.
- Default individual expanded entry limit: 64 MiB.
- Default aggregate expanded ZIP limit: 512 MiB.
- XML and binary-record recursion is bounded to 256 levels.
- Encrypted packages fail closed; password decryption is not attempted.
- External relationships are not fetched by the Rust parser.
- Hyperlinks are restricted to HTTP(S), mail, telephone, relative, and fragment targets and render
  with `noopener noreferrer`; active schemes such as `javascript:` and `data:` are rejected.
- Object URLs and asynchronous slide resources are revoked on teardown.

Consumers handling hostile uploads should also enforce request-size limits before the browser,
serve Wasm with `application/wasm`, and use a restrictive CSP. If legacy EMF/WMF previews are
enabled, the CSP must also permit the resulting browser image URLs.
