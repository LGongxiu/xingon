# XINGON!

**XINGON!** es un juego de cartas lleno de faroles, creado por **Laia G. Grasas**. En tu turno dices qué cartas pones, y los demás deciden si te creen o te pillan.

🌐 **Web del juego:** https://lgongxiu.github.io/xingon/

La web reúne todo lo necesario para aprender a jugar y está disponible en **castellano, català i English**.

- 🎬 Vídeo de cómo se juega, uno por idioma
- 📖 Instrucciones completas, resumen en 30 segundos y variante fácil
- ❓ Preguntas frecuentes
- 🧮 Marcador de puntos para usar desde el móvil durante la partida
- 🃏 La baraja completa, con filtros por color
- ✍️ Sobre la autora

## El juego

| | |
|---|---|
| Jugadores | 2–8 |
| Edad | +8 años |
| Duración | < 20 minutos |
| Contenido | 68 cartas: 60 numeradas (4 colores × 3 figuras × 5 números), 5 comodines y 3 especiales |

Cada color tiene su propio símbolo en las esquinas (★ amarillo, ☾ azul, árbol verde, ✚ rojo), para que las cartas se distingan aunque no se vean bien los colores.

## Estructura del repositorio

```
index.html            La web completa: diseño, textos de los 3 idiomas y funciones
assets/
├── logo.png          Logo de XINGON!
├── favicon.png       Icono de la pestaña del navegador
├── apple-touch-icon.png
├── og.jpg            Imagen que aparece al compartir el enlace
├── video_es.mp4      Vídeo en castellano
├── video_ca.mp4      Vídeo en catalán
├── video_en.mp4      Vídeo en inglés
├── poster_*.jpg      Portada de cada vídeo
├── *.webp            Cartas usadas en la página
└── deck/             Las 68 cartas de la galería
```

## Cómo editar los textos

Todos los textos están en `index.html`, dentro de la parte `<script>` del final, separados por idioma (`"es"`, `"ca"` y `"en"`). Cada texto tiene un nombre, por ejemplo:

```js
"tagline": "Un juego de cartas lleno de faroles: ...",
"goal": "Sé el primero en quedarte sin cartas.",
```

Para cambiar un texto, edita solo lo que hay entre comillas después de los dos puntos. Recuerda cambiarlo en los tres idiomas.

## Cómo actualizar la web

1. Edita o sube los archivos en GitHub y pulsa **Commit changes**.
2. GitHub Pages publica los cambios solo. Puedes seguir el proceso en la pestaña **Actions**; tarda 1–2 minutos.
3. Recarga la web sin caché (`Ctrl + F5` en Windows o `Cmd + Shift + R` en Mac) para ver la versión nueva.

## Derechos

© Laia G. Grasas. Todos los derechos reservados.

El juego XINGON!, su nombre, logo, diseño de cartas, instrucciones y vídeos son obra de la autora. Este repositorio se publica para alojar la web del juego. No está permitido copiar, reproducir ni comercializar el juego o sus materiales sin permiso.
