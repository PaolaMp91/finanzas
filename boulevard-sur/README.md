# Boulevard Sur · Club Residencial — Landing cinemática "3D Scroll"

Página web cinemática scroll-driven para **Boulevard Sur** (Zona 8 de Mixco, Ciudad San Cristóbal, Guatemala).

## Cómo verla en localhost

```bash
cd boulevard-sur
python3 -m http.server 8080
# abrir http://localhost:8080
```

> Se necesita un servidor local (no abrir el archivo con doble clic) para que los videos
> se puedan "scrubbear" con el scroll correctamente.

## Experiencia

| Acto | Qué pasa al hacer scroll |
|---|---|
| Hero | **"VIVE AL RITMO DEL SUR — Bienvenido a Boulevard Sur"** sobre el edificio cubierto por la manta |
| 01 · La manta cae | El clip del edificio cubierto se reproduce con el scroll, la manta digital se levanta pliegue a pliegue y un destello revela el edificio |
| 02 · Familia | Clip cinemático de familia con movimientos de cámara rítmicos y sorpresivos |
| 03 · Interior → Piscina | Recorrido del apartamento y golpe de cámara (whip-pan) hacia la piscina del club |
| Proyecto | Stats animados, amenidades, ubicación y distancias reales del brochure |
| Plantas | Tipologías de 2 y 3 habitaciones con plano esquemático (52.39 m² / 64.69 m²) |
| Cotizador | Enganche 10% en cuotas + crédito bancario (tasa ajustable, 15/20/25 años) |

## Identidad de marca

Paleta y tipografías tomadas del **B SUR Brandbook** oficial (Teams › MERCADEO › 02. BOULEVARD SUR › MANUAL DE MARCA):

- Verde `#577E2E` · Azul `#184C61` · Blanco `#FFFFFF` · Gris `#666666`
- Acento dorado `#C9A96A` agregado para el look de lujo
- Títulos: Archivo (sustituto web de *Rele*) · Párrafos: Manrope (sustituto web de *Avenir Next*)

## Assets generados con IA (Runway)

- `assets/img/*.png` — renders cinematográficos generados con **nano-banana-pro** (basados en los renders del proyecto en Teams › RENDERS › IA)
- `assets/video/clip1-reveal.mp4` — edificio cubierto, la manta cae (gen-4-turbo, 5 s)
- `assets/video/clip2-familia.mp4` — familia close-up con cámara rítmica (gen-4-turbo, 5 s)
- `assets/video/clip3-interior-piscina.mp4` — interior → piscina (pendiente: el workspace de Runway alcanzó su límite diario; mientras tanto la transición se hace por código con las imágenes). Para agregarlo: correr el workflow **Fetch Boulevard Sur media assets** con la URL del clip en el input `clip3_url`.

Los archivos de media se descargan al repo con el workflow de GitHub Actions
`.github/workflows/fetch-assets.yml` (los CDNs de Runway expiran las URLs firmadas).

## Datos del proyecto (fuentes)

- Brochure oficial `BS Brochure 2025may` (Teams · MERCADEO)
- boulevardsur.gt / prensa: precios Q550,000–Q750,000, ~600 apartamentos, entrega 2027
- Dirección: 6a avenida 22-80, Zona 8 de Mixco, San Cristóbal
