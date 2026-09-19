# Integridad de lo vendorizado

Registro para que una sesión futura pueda probar que el archivo del repo es el que publicó su
autor, sin confiar en el nombre (amarillo 7 del auditor-repo, 2026-09-11).

| Archivo | Versión | Origen | sha256 | Verificado |
|---|---|---|---|---|
| `msal-browser.min.js` | @azure/msal-browser 4.29.0 | `https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.29.0/lib/msal-browser.min.js` | `d822083e23e729bd49248c54b68c51b6d5dbcff276a5ab3ab46f57b295de9cb7` | 2026-09-11, descarga fresca del origen = mismo hash |

Cómo re-verificar (desde la raíz de `proyectos/`):

```
sha256sum calytek-planta-app/app/vendor/msal-browser.min.js
curl -sL https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.29.0/lib/msal-browser.min.js | sha256sum
```

Los dos hashes deben coincidir entre sí y con la tabla. Al subir de versión: nueva fila, nunca
sobrescribir la anterior (la vieja prueba qué se corrió hasta esa fecha).

**`integrity=` de `index.html` (S-03, 2026-09-19):** el `<script>` lleva el sha384 del mismo archivo,
`sha384-8rQAl229+briV9tfGmSO/NUMdiZrVcD/TzWsAcIt/NsxuH7yYizuHByfyygUVwVe` (`openssl dgst -sha384
-binary vendor/msal-browser.min.js | base64`). `test/vendor.test.js` (en `npm test`) coteja el sha256 de
la ÚLTIMA fila de esta tabla y el `integrity=` contra el archivo del repo: subir el vendor obliga a
tocar los tres a la vez. **Cierre de versión (S-05):** `npm view @azure/msal-browser version` contra la
fila vigente; se sube cuando el changelog de la versión nueva traiga «security».
