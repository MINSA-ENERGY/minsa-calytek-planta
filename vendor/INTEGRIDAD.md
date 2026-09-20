# Integridad de lo vendorizado

Registro para que una sesión futura pueda probar que el archivo del repo es el que publicó su
autor, sin confiar en el nombre (amarillo 7 del auditor-repo, 2026-09-11).

| Archivo | Versión | Origen | sha256 | Verificado |
|---|---|---|---|---|
| `msal-browser.min.js` | @azure/msal-browser 4.29.0 | `https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.29.0/lib/msal-browser.min.js` | `d822083e23e729bd49248c54b68c51b6d5dbcff276a5ab3ab46f57b295de9cb7` | 2026-09-11, descarga fresca del origen = mismo hash |
| `msal-browser.min.js` | @azure/msal-browser 5.22.0 | `https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.22.0/lib/msal-browser.min.js` | `5ce42b98842c06a0d00233253f46684f43dd398b46cb7dd5e1441f1bfd9caa6d` | 2026-09-20, dos descargas del origen + `lib/msal-browser.min.js` del tarball de npm (`registry.npmjs.org/@azure/msal-browser/-/msal-browser-5.22.0.tgz`) = mismo hash |

## Fuentes (S-17, 2026-09-20)

Las 8 `woff2` de `vendor/fuentes/` (subconjunto **latin** que Google Fonts sirve a un Chrome moderno; las
mismas que `minsa-proyectos-app`, byte a byte). Una fuente no ejecuta script, pero es un binario que parsea
el navegador y `@font-face` no admite `integrity=`: el hash aquí y `test/vendor.test.js` son la única prueba
de que son las publicadas. Verificadas el 2026-09-20 contra descarga fresca del origen (mismo hash las 8).

| Archivo | Versión | Origen | sha256 |
|---|---|---|---|
| `fuentes/Saira-500.woff2` | Saira v23 (variable, wght 100-900; Google sirve el MISMO archivo para 500 y 600) | `https://fonts.gstatic.com/s/saira/v23/memjYa2wxmKQyPMrZX79wwYZQMhsyuSLiIvS.woff2` | `7eb811eb14b2ee22e3fba942b25c6cd062ff050bde10d29af1a4e16f99712e17` |
| `fuentes/Saira-600.woff2` | Saira v23 (el mismo archivo que Saira-500) | `https://fonts.gstatic.com/s/saira/v23/memjYa2wxmKQyPMrZX79wwYZQMhsyuSLiIvS.woff2` | `7eb811eb14b2ee22e3fba942b25c6cd062ff050bde10d29af1a4e16f99712e17` |
| `fuentes/Barlow-400.woff2` | Barlow v13 | `https://fonts.gstatic.com/s/barlow/v13/7cHpv4kjgoGqM7E_DMs5.woff2` | `b0a8ad37ac45f5fb22ced461576db72e44e295107aad7a9c8a7a4bad728fd03b` |
| `fuentes/Barlow-500.woff2` | Barlow v13 | `https://fonts.gstatic.com/s/barlow/v13/7cHqv4kjgoGqM7E3_-gs51os.woff2` | `cd759df8ef9efc98fee14307b4eb5ba27f08b1f8f2f3ad2872432e25c89907a8` |
| `fuentes/Barlow-600.woff2` | Barlow v13 | `https://fonts.gstatic.com/s/barlow/v13/7cHqv4kjgoGqM7E30-8s51os.woff2` | `4b52ddd4836b592df0e4832b8286956883cdc651b015126bdd18f184b7f90cc3` |
| `fuentes/IBMPlexMono-400.woff2` | IBM Plex Mono v20 | `https://fonts.gstatic.com/s/ibmplexmono/v20/-F63fjptAgt5VM-kVkqdyU8n1i8q1w.woff2` | `08949f728dc52d528e69b1667d15c89a5686a4ee9a296ff90983985f99c380f7` |
| `fuentes/IBMPlexMono-500.woff2` | IBM Plex Mono v20 | `https://fonts.gstatic.com/s/ibmplexmono/v20/-F6qfjptAgt5VM-kVkqdyU8n3twJwlBFgg.woff2` | `01d285447409c8a588692162439a038b8cbd7871309ee20267b0d2d91c6e8e22` |
| `fuentes/BaiJamjuree-700.woff2` | Bai Jamjuree v13 | `https://fonts.gstatic.com/s/baijamjuree/v13/LDIqapSCOBt_aeQQ7ftydoa05efelJo0.woff2` | `01f179f3f3b1ae29e89409cfdc4665af5c775995dfb2c13f6c8f5d1907028119` |

Cómo re-verificar: `curl -sL -A "Mozilla/5.0 (Windows NT 10.0) Chrome/128" "https://fonts.googleapis.com/css2?family=Saira:wght@500;600&family=Barlow:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&family=Bai+Jamjuree:wght@700"`
da los `@font-face`; el bloque con `unicode-range: U+0000-00FF` de cada familia/peso trae la URL de arriba (la URL cambia
cuando Google sube la versión de la fuente: entonces se re-verifica y se agrega fila, sin borrar la vieja).

Cómo re-verificar `msal-browser.min.js` (desde la raíz de `proyectos/`):

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
