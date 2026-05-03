# Éditeur Markdown WYSIWYG

Éditeur Markdown en React avec aperçu en temps réel, barre d’outils de mise en forme et export des contenus en HTML, Word et PDF.

## Fonctionnalités

- Édition Markdown avec prévisualisation live côte à côte.
- Barre d’outils pour insérer rapidement titres, listes, citations, liens, blocs de code, tableaux et images.
- Raccourcis clavier pour le gras, l’italique, le code inline et les liens.
- Compteurs de mots et de caractères affichés en temps réel.
- Export du contenu au format HTML, DOCX et PDF.
- Rendu Markdown enrichi avec support GFM et surlignage du code.

## Technologies

- React 18
- TypeScript
- Vite
- react-markdown
- remark-gfm
- rehype-highlight
- docx
- html2canvas
- jsPDF

## Prérequis

- Node.js 18 ou supérieur
- npm

## Installation

```bash
npm install
```

## Démarrage en local

```bash
npm run dev
```

Vite affiche alors l’adresse locale de l’application dans le terminal.

## Build de production

```bash
npm run build
```

## Aperçu du projet

L’application est organisée autour d’un éditeur Markdown à gauche et d’une prévisualisation à droite. Le contenu initial sert de démonstration avec des titres, un tableau, une image et un bloc de code pour montrer le rendu disponible dès le chargement.

## Scripts

- `npm run dev` : lance le serveur de développement.
- `npm run build` : génère la version de production.
- `npm run preview` : prévisualise le build localement.
