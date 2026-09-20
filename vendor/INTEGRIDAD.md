# Integridad de lo vendorizado

Registro para que una sesión futura pueda probar que el archivo del repo es el que publicó su
autor, sin confiar en el nombre (amarillo 7 del auditor-repo, 2026-09-11).

| Archivo | Versión | Origen | sha256 | Verificado |
|---|---|---|---|---|
| `msal-browser.min.js` | @azure/msal-browser 4.29.0 | `https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.29.0/lib/msal-browser.min.js` | `d822083e23e729bd49248c54b68c51b6d5dbcff276a5ab3ab46f57b295de9cb7` | 2026-09-11, descarga fresca del origen = mismo hash |
| `msal-browser.min.js` | @azure/msal-browser 5.22.0 | `https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.22.0/lib/msal-browser.min.js` | `5ce42b98842c06a0d00233253f46684f43dd398b46cb7dd5e1441f1bfd9caa6d` | 2026-09-20, dos descargas del origen + `lib/msal-browser.min.js` del tarball de npm (`registry.npmjs.org/@azure/msal-browser/-/msal-browser-5.22.0.tgz`) = mismo hash |

Cómo re-verificar (desde la raíz de `proyectos/`):

```
sha256sum calytek-planta-app/app/vendor/msal-browser.min.js
curl -sL https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.22.0/lib/msal-browser.min.js | sha256sum
```

Los dos hashes deben coincidir entre sí y con la tabla. Al subir de versión: nueva fila, nunca
sobrescribir la anterior (la vieja prueba qué se corrió hasta esa fecha).

**`integrity=` de `index.html` (S-03, 2026-09-19):** el `<script>` lleva el sha384 del mismo archivo,
`sha384-0xw/kzSK+WLDaLIkXwqFOXYqCnxt2agAhE5d3NN2ynU8GW5pry4+6dDtv4U2Fhkx` (`openssl dgst -sha384
-binary vendor/msal-browser.min.js | base64`). `test/vendor.test.js` (en `npm test`) coteja el sha256 de
la ÚLTIMA fila de esta tabla y el `integrity=` contra el archivo del repo, y desde la v0.30.0 (S-13) carga
el bundle REAL en Node y exige que `msal.PublicClientApplication` exponga los 7 métodos que usa `app.js`
(`initialize`, `getAllAccounts`, `acquireTokenSilent`, `acquireTokenRedirect`, `handleRedirectPromise`,
`loginRedirect`, `logoutRedirect`), `InteractionRequiredAuthError` y `CacheLookupPolicy`: subir el vendor
obliga a tocar los tres a la vez, y un bundle que cambie la API falla en `npm test` aunque la E2E (que usa un
MSAL falso) salga verde. **Cierre de versión (S-05, reescrito en S-13):** `npm view @azure/msal-browser
version` contra la fila vigente; se compara contra la **última 5.x** (la línea 4.x terminó en 4.30.0 el
2026-03-18 y no recibe parches) y se sube cuando la 5.x nueva traiga «security» en el changelog o cuando la
vigente lleve más de 6 meses sin subir. **Lo que la E2E no prueba:** el login real; tras cada subida Carlos
entra una vez desde escritorio y una desde el celular, y si falla, revertir es restaurar
`vendor/msal-browser.min.js` + `integrity=` + esta fila desde el commit anterior.
