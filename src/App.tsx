import { useEffect, useMemo, useRef, useState } from "react";
import ReactDOMServer from "react-dom/server";
import {
  AlignmentType,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import ReactMarkdown from "react-markdown";
import rehypeHighlight from "rehype-highlight";
import remarkGfm from "remark-gfm";

const INITIAL_MARKDOWN = `# Éditeur Markdown

Écris à gauche, visualise à droite.

## Raccourcis

- **Ctrl/Cmd + B** : gras
- **Ctrl/Cmd + I** : italique
- **Ctrl/Cmd + K** : lien
- **Ctrl/Cmd + \`** : code inline

## Tableau

| Fonction | Statut |
| --- | --- |
| Édition | Active |
| Prévisualisation | En direct |

## Image

![Bannière](https://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&w=1200&q=80)

\`\`\`ts
const message: string = 'Hello Markdown';
console.log(message);
\`\`\`
`;

const THEME_STORAGE_KEY = "markdown-editor-theme";
type ThemeMode = "light" | "dark";

function getPreferredTheme(): ThemeMode {
  if (typeof window === "undefined") {
    return "light";
  }

  const savedTheme = window.localStorage.getItem(THEME_STORAGE_KEY);
  if (savedTheme === "light" || savedTheme === "dark") {
    return savedTheme;
  }

  return window.matchMedia("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light";
}

function wrapSelection(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  wrapper: string,
) {
  const selectedText = text.slice(selectionStart, selectionEnd);
  const nextText = `${text.slice(0, selectionStart)}${wrapper}${selectedText}${wrapper}${text.slice(selectionEnd)}`;

  return {
    text: nextText,
    selectionStart: selectionStart + wrapper.length,
    selectionEnd: selectionEnd + wrapper.length,
  };
}

type SelectionRange = {
  text: string;
  selectionStart: number;
  selectionEnd: number;
};

function insertBlock(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  block: string,
): SelectionRange {
  const hasLeadingBreak =
    selectionStart > 0 && text[selectionStart - 1] !== "\n";
  const hasTrailingBreak =
    selectionEnd < text.length && text[selectionEnd] !== "\n";
  const replacement = `${hasLeadingBreak ? "\n" : ""}${block}${hasTrailingBreak ? "\n" : ""}`;

  return replaceSelection(
    text,
    selectionStart,
    selectionEnd,
    replacement,
    replacement.length,
    replacement.length,
  );
}

function applyLinePrefix(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  prefix: string,
): SelectionRange {
  const startOfLine = text.lastIndexOf("\n", selectionStart - 1) + 1;
  const selectedText = text.slice(startOfLine, selectionEnd);
  const replacement = selectedText
    .split("\n")
    .map((line) => `${prefix}${line}`)
    .join("\n");

  return replaceSelection(
    text,
    startOfLine,
    selectionEnd,
    replacement,
    replacement.length,
    replacement.length,
  );
}

function createTableTemplate() {
  return [
    "| Colonne 1 | Colonne 2 | Colonne 3 |",
    "| --- | --- | --- |",
    "| Donnée 1 | Donnée 2 | Donnée 3 |",
    "| Donnée 4 | Donnée 5 | Donnée 6 |",
  ].join("\n");
}

function createCodeBlockTemplate(language: string) {
  return `\`\`\`${language}\n// Votre code ici\n\`\`\``;
}

function insertAroundSelection(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  before: string,
  after: string,
) {
  const selectedText = text.slice(selectionStart, selectionEnd);
  const nextText = `${text.slice(0, selectionStart)}${before}${selectedText}${after}${text.slice(selectionEnd)}`;

  return {
    text: nextText,
    selectionStart: selectionStart + before.length,
    selectionEnd: selectionEnd + before.length,
  };
}

function replaceSelection(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  replacement: string,
  selectStartOffset = replacement.length,
  selectEndOffset = replacement.length,
) {
  const nextText = `${text.slice(0, selectionStart)}${replacement}${text.slice(selectionEnd)}`;

  return {
    text: nextText,
    selectionStart: selectionStart + selectStartOffset,
    selectionEnd: selectionStart + selectEndOffset,
  };
}

function MarkdownContent({ markdown }: { markdown: string }) {
  return (
    <ReactMarkdown
      remarkPlugins={[remarkGfm]}
      rehypePlugins={[rehypeHighlight]}
    >
      {markdown}
    </ReactMarkdown>
  );
}

function getExportPageStyle(theme: ThemeMode) {
  const isDark = theme === "dark";
  const pageBackground = isDark ? "#020617" : "#f8fafc";
  const pageAccent = isDark
    ? "rgba(56, 189, 248, 0.12)"
    : "rgba(59, 130, 246, 0.12)";
  const textPrimary = isDark ? "#e2e8f0" : "#111827";
  const textSecondary = isDark ? "#cbd5e1" : "#475569";
  const surface = isDark ? "rgba(15, 23, 42, 0.92)" : "white";
  const surfaceBorder = isDark ? "rgba(148, 163, 184, 0.18)" : "#e2e8f0";
  const codeBackground = isDark ? "#020617" : "#0f172a";
  const codeForeground = "#e2e8f0";

  return `
  :root {
    color-scheme: ${theme};
  }
  body {
    font-family: Inter, system-ui, sans-serif;
    line-height: 1.7;
    margin: 0;
    padding: 2rem;
    color: ${textPrimary};
    background:
      radial-gradient(circle at top left, ${pageAccent}, transparent 30%),
      ${pageBackground};
  }
  .content {
    max-width: 900px;
    margin: 0 auto;
    background: ${surface};
    color: ${textPrimary};
    padding: 2rem;
    border-radius: 1rem;
    border: 1px solid ${surfaceBorder};
    box-shadow: 0 20px 60px rgba(15, 23, 42, 0.08);
  }
  h1, h2, h3, h4, h5, h6 { line-height: 1.2; }
  pre {
    padding: 1rem;
    overflow: auto;
    border-radius: 0.75rem;
    background: ${codeBackground};
    color: ${codeForeground};
  }
  code { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; }
  blockquote {
    border-left: 4px solid ${surfaceBorder};
    padding-left: 1rem;
    color: ${textSecondary};
    margin-left: 0;
  }
  table {
    width: 100%;
    border-collapse: collapse;
  }
  th, td {
    border: 1px solid ${surfaceBorder};
    padding: 0.75rem;
  }
  a {
    color: ${isDark ? "#60a5fa" : "#2563eb"};
  }
  img {
    max-width: 100%;
    border-radius: 0.75rem;
  }
`;
}

function buildExportDocumentHtml(contentHtml: string, theme: ThemeMode) {
  return `<!doctype html>
<html lang="fr">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Export Markdown</title>
    <style>${getExportPageStyle(theme)}</style>
  </head>
  <body data-theme="${theme}">
    <main class="content">
      ${contentHtml}
    </main>
  </body>
</html>`;
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

type InlineStyleState = {
  bold?: boolean;
  italics?: boolean;
  code?: boolean;
};

function createWordRunFromNode(
  node: Node,
  styles: InlineStyleState = {},
): Array<TextRun | ExternalHyperlink> {
  if (node.nodeType === Node.TEXT_NODE) {
    const value = node.textContent ?? "";
    return value
      ? [
          new TextRun({
            text: value,
            bold: styles.bold,
            italics: styles.italics,
            font: styles.code ? "Courier New" : undefined,
            size: styles.code ? 20 : undefined,
          }),
        ]
      : [];
  }

  if (!(node instanceof HTMLElement)) {
    return [];
  }

  if (node.tagName === "BR") {
    return [new TextRun({ text: "", break: 1 })];
  }

  if (node.tagName === "STRONG" || node.tagName === "B") {
    return Array.from(node.childNodes).flatMap((child) => {
      return createWordRunFromNode(child, { ...styles, bold: true });
    });
  }

  if (node.tagName === "EM" || node.tagName === "I") {
    return Array.from(node.childNodes).flatMap((child) => {
      return createWordRunFromNode(child, { ...styles, italics: true });
    });
  }

  if (node.tagName === "CODE") {
    return [
      new TextRun({
        text: node.textContent ?? "",
        font: "Courier New",
        size: 20,
        bold: styles.bold,
        italics: styles.italics,
      }),
    ];
  }

  if (node.tagName === "A") {
    const href = node.getAttribute("href") ?? "";
    const text = node.textContent?.trim() || href;
    return [
      new ExternalHyperlink({
        link: href,
        children: [
          new TextRun({
            text,
            color: "0563C1",
            underline: { type: "single" },
          }),
        ],
      }),
    ];
  }

  return Array.from(node.childNodes).flatMap((child) =>
    createWordRunFromNode(child, styles),
  );
}

function createDocxParagraphFromElement(element: Element): Paragraph {
  const tagName = element.tagName.toLowerCase();

  if (tagName === "h1") {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { after: 180 },
      children: createWordRunFromNode(element),
    });
  }

  if (tagName === "h2") {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { after: 160 },
      children: createWordRunFromNode(element),
    });
  }

  if (tagName === "h3") {
    return new Paragraph({
      heading: HeadingLevel.HEADING_3,
      spacing: { after: 140 },
      children: createWordRunFromNode(element),
    });
  }

  if (tagName === "blockquote") {
    return new Paragraph({
      children: createWordRunFromNode(element),
      indent: { left: 420 },
      spacing: { after: 120 },
    });
  }

  if (tagName === "pre") {
    return new Paragraph({
      children: [
        new TextRun({
          text: element.textContent ?? "",
          font: "Courier New",
        }),
      ],
      shading: { type: "clear", color: "auto", fill: "F8FAFC" },
      border: {
        left: { style: "single", size: 8, color: "CBD5E1" },
        top: { style: "single", size: 6, color: "E2E8F0" },
        right: { style: "single", size: 6, color: "E2E8F0" },
        bottom: { style: "single", size: 6, color: "E2E8F0" },
      },
      spacing: { before: 120, after: 120 },
    });
  }

  if (tagName === "p") {
    return new Paragraph({
      children: createWordRunFromNode(element),
      spacing: { after: 120 },
    });
  }

  if (tagName === "ul" || tagName === "ol") {
    const children: Paragraph[] = [];
    Array.from(element.children).forEach((listItem, index) => {
      const prefix = tagName === "ol" ? `${index + 1}. ` : "• ";
      children.push(
        new Paragraph({
          children: [
            new TextRun({ text: prefix, bold: true }),
            ...createWordRunFromNode(listItem),
          ],
          spacing: { after: 80 },
        }),
      );
    });
    return children[0] ?? new Paragraph("");
  }

  if (tagName === "hr") {
    return new Paragraph({
      text: "────────────────────────",
      spacing: { after: 120 },
    });
  }

  if (tagName === "img") {
    const alt = element.getAttribute("alt") ?? "Image";
    const src = element.getAttribute("src") ?? "";
    return new Paragraph({
      children: [new TextRun({ text: `${alt} (${src})`, italics: true })],
      spacing: { after: 120 },
    });
  }

  return new Paragraph({
    children: createWordRunFromNode(element),
    spacing: { after: 120 },
  });
}

function createDocxTableFromElement(element: HTMLTableElement) {
  const rows = Array.from(element.querySelectorAll("tr")).map((row) => {
    const cells = Array.from(row.children).map(
      (cell) =>
        new TableCell({
          children: [
            new Paragraph({
              children: [new TextRun(cell.textContent ?? "")],
            }),
          ],
        }),
    );

    return new TableRow({ children: cells });
  });

  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
}

function buildWordDocumentFromHtml(contentHtml: string) {
  const parser = new DOMParser();
  const documentNode = parser.parseFromString(
    `<main>${contentHtml}</main>`,
    "text/html",
  );
  const root = documentNode.querySelector("main");

  const blocks: Array<Paragraph | Table> = [];
  if (!root) {
    return blocks;
  }

  Array.from(root.children).forEach((child) => {
    if (child.tagName.toLowerCase() === "table") {
      blocks.push(createDocxTableFromElement(child as HTMLTableElement));
      return;
    }

    blocks.push(createDocxParagraphFromElement(child));
  });

  return blocks;
}

function createPdfSurface(contentHtml: string, theme: ThemeMode) {
  const wrapper = document.createElement("div");
  wrapper.style.position = "fixed";
  wrapper.style.left = "-10000px";
  wrapper.style.top = "0";
  wrapper.style.width = "900px";
  wrapper.innerHTML = `<style>${getExportPageStyle(theme)}</style><main class="content">${contentHtml}</main>`;
  document.body.appendChild(wrapper);

  return wrapper;
}

export default function App() {
  const [markdown, setMarkdown] = useState(INITIAL_MARKDOWN);
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const [characterCount, setCharacterCount] = useState(INITIAL_MARKDOWN.length);
  const [theme, setTheme] = useState<ThemeMode>(getPreferredTheme);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    window.localStorage.setItem(THEME_STORAGE_KEY, theme);
  }, [theme]);

  const handleChange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    const nextValue = event.target.value;
    setMarkdown(nextValue);
    setCharacterCount(nextValue.length);
  };

  const updateEditor = (next: SelectionRange) => {
    setMarkdown(next.text);
    setCharacterCount(next.text.length);

    window.requestAnimationFrame(() => {
      const textarea = textareaRef.current;
      if (!textarea) {
        return;
      }

      textarea.focus();
      textarea.setSelectionRange(
        next.selectionStart,
        next.selectionEnd,
        "forward",
      );
    });
  };

  const insertText = (
    textToInsert: string,
    selectStartOffset?: number,
    selectEndOffset?: number,
  ) => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const { selectionStart, selectionEnd, value } = textarea;
    updateEditor(
      replaceSelection(
        value,
        selectionStart,
        selectionEnd,
        textToInsert,
        selectStartOffset ?? textToInsert.length,
        selectEndOffset ?? textToInsert.length,
      ),
    );
  };

  const insertHeading = (level: 1 | 2 | 3) => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const prefix = `${"#".repeat(level)} `;
    const { selectionStart, selectionEnd, value } = textarea;
    updateEditor(applyLinePrefix(value, selectionStart, selectionEnd, prefix));
  };

  const insertBlockquote = () => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const { selectionStart, selectionEnd, value } = textarea;
    updateEditor(applyLinePrefix(value, selectionStart, selectionEnd, "> "));
  };

  const insertList = (ordered: boolean) => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const { selectionStart, selectionEnd, value } = textarea;
    updateEditor(
      applyLinePrefix(
        value,
        selectionStart,
        selectionEnd,
        ordered ? "1. " : "- ",
      ),
    );
  };

  const insertLink = () => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const href = window.prompt("URL du lien", "https://") ?? "https://";
    const { selectionStart, selectionEnd, value } = textarea;
    const selectedText =
      value.slice(selectionStart, selectionEnd) || "texte du lien";
    const replacement = `[${selectedText}](${href})`;
    updateEditor(
      replaceSelection(
        value,
        selectionStart,
        selectionEnd,
        replacement,
        1,
        1 + selectedText.length,
      ),
    );
  };

  const insertCodeBlock = () => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const language = window.prompt("Langage du bloc de code", "ts") ?? "txt";
    const { selectionStart, selectionEnd, value } = textarea;
    updateEditor(
      insertBlock(
        value,
        selectionStart,
        selectionEnd,
        createCodeBlockTemplate(language),
      ),
    );
  };

  const insertTable = () => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const { selectionStart, selectionEnd, value } = textarea;
    updateEditor(
      insertBlock(value, selectionStart, selectionEnd, createTableTemplate()),
    );
  };

  const promptForImage = () => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const url = window.prompt(
      "URL de l'image",
      "https://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&w=1200&q=80",
    );

    if (!url) {
      return;
    }

    const alt =
      window.prompt("Texte alternatif", "Illustration") ?? "Illustration";
    const { selectionStart, selectionEnd, value } = textarea;
    updateEditor(
      insertBlock(value, selectionStart, selectionEnd, `![${alt}](${url})`),
    );
  };

  const applyShortcut = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const isMac = navigator.platform.toLowerCase().includes("mac");
    const commandKey = isMac ? event.metaKey : event.ctrlKey;

    if (!commandKey) {
      return;
    }

    const { selectionStart, selectionEnd, value } = textarea;
    if (event.key.toLowerCase() === "b") {
      event.preventDefault();
      updateEditor(wrapSelection(value, selectionStart, selectionEnd, "**"));
      return;
    }

    if (event.key.toLowerCase() === "i") {
      event.preventDefault();
      updateEditor(wrapSelection(value, selectionStart, selectionEnd, "_"));
      return;
    }

    if (event.key === "`") {
      event.preventDefault();
      updateEditor(wrapSelection(value, selectionStart, selectionEnd, "`"));
      return;
    }

    if (event.key.toLowerCase() === "k") {
      event.preventDefault();
      insertLink();
      return;
    }
  };

  const exportHtml = () => {
    const contentHtml = ReactDOMServer.renderToStaticMarkup(
      <MarkdownContent markdown={markdown} />,
    );
    const documentHtml = buildExportDocumentHtml(contentHtml, theme);

    downloadBlob(
      new Blob([documentHtml], { type: "text/html;charset=utf-8" }),
      "markdown-export.html",
    );
  };

  const exportWord = async () => {
    const contentHtml = ReactDOMServer.renderToStaticMarkup(
      <MarkdownContent markdown={markdown} />,
    );
    const blocks = buildWordDocumentFromHtml(contentHtml);
    const document = new Document({
      sections: [
        {
          properties: {},
          children:
            blocks.length > 0 ? blocks : [new Paragraph("Document vide")],
        },
      ],
    });

    const blob = await Packer.toBlob(document);
    downloadBlob(blob, "markdown-export.docx");
  };

  const exportPdf = async () => {
    const contentHtml = ReactDOMServer.renderToStaticMarkup(
      <MarkdownContent markdown={markdown} />,
    );
    const wrapper = createPdfSurface(contentHtml, theme);

    try {
      const pdf = new jsPDF("p", "pt", "a4");
      const contentElement = wrapper.querySelector(
        ".content",
      ) as HTMLElement | null;

      if (!contentElement) {
        return;
      }

      const canvas = await html2canvas(contentElement, {
        scale: 2,
        backgroundColor: "#f8fafc",
        useCORS: true,
      });

      const imgData = canvas.toDataURL("image/png");
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const imgWidth = pageWidth;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      let heightLeft = imgHeight;
      let position = 0;

      pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
      heightLeft -= pageHeight;

      while (heightLeft > 0) {
        position = heightLeft - imgHeight;
        pdf.addPage();
        pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;
      }

      pdf.save("markdown-export.pdf");
    } finally {
      wrapper.remove();
    }
  };

  const preview = useMemo(() => markdown, [markdown]);
  const wordCount = useMemo(
    () => markdown.trim().split(/\s+/).filter(Boolean).length,
    [markdown],
  );

  const themeLabel = theme === "dark" ? "Mode clair" : "Mode sombre";

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-copy-block">
          <p className="eyebrow">React + TypeScript + Vite</p>
          <h1>Éditeur Markdown avec aperçu en temps réel</h1>
          <p className="hero-copy">
            Saisie à gauche, rendu propre à droite, raccourcis clavier,
            insertion de tableaux et d'images, puis export HTML.
          </p>
          <div className="hero-stats">
            <span>{wordCount} mots</span>
            <span>{characterCount} caractères</span>
            <span>Grille en direct</span>
          </div>
        </div>
        <div className="hero-actions">
          <button
            className="theme-toggle"
            type="button"
            onClick={() => setTheme(theme === "dark" ? "light" : "dark")}
            aria-pressed={theme === "dark"}
          >
            {themeLabel}
          </button>
          <div className="export-group" aria-label="Exports">
            <button
              className="export-button"
              type="button"
              onClick={exportHtml}
            >
              HTML
            </button>
            <button
              className="export-button export-button--secondary"
              type="button"
              onClick={exportWord}
            >
              Word
            </button>
            <button
              className="export-button export-button--secondary"
              type="button"
              onClick={exportPdf}
            >
              PDF
            </button>
          </div>
        </div>
      </header>

      <main className="workspace">
        <section className="panel editor-panel" aria-label="Zone d'édition">
          <div className="panel-header panel-header--stacked">
            <div className="panel-heading-row">
              <h2>Markdown</h2>
              <span>Ctrl/Cmd + B, I, K, `</span>
            </div>
            <div
              className="toolbar"
              role="toolbar"
              aria-label="Barre d'outils Markdown"
            >
              <button type="button" onClick={() => insertHeading(1)}>
                H1
              </button>
              <button type="button" onClick={() => insertHeading(2)}>
                H2
              </button>
              <button type="button" onClick={() => insertHeading(3)}>
                H3
              </button>
              <button type="button" onClick={() => insertBlockquote()}>
                &gt;
              </button>
              <button type="button" onClick={() => insertList(false)}>
                Liste
              </button>
              <button type="button" onClick={() => insertList(true)}>
                Num.
              </button>
              <button type="button" onClick={() => insertLink()}>
                Lien
              </button>
              <button type="button" onClick={() => insertCodeBlock()}>
                Code
              </button>
              <button type="button" onClick={() => insertTable()}>
                Tableau
              </button>
              <button type="button" onClick={() => promptForImage()}>
                Image
              </button>
              <button
                type="button"
                onClick={() => insertText("**gras**", 2, 6)}
              >
                Gras
              </button>
              <button
                type="button"
                onClick={() => insertText("_italique_", 1, 9)}
              >
                Italique
              </button>
            </div>
          </div>
          <textarea
            ref={textareaRef}
            value={markdown}
            onChange={handleChange}
            onKeyDown={applyShortcut}
            spellCheck={false}
            aria-label="Éditeur Markdown"
          />
        </section>

        <section className="panel preview-panel" aria-label="Prévisualisation">
          <div className="panel-header">
            <h2>Prévisualisation</h2>
            <span>Rendu live</span>
          </div>
          <article className="preview-content">
            <ReactMarkdown
              remarkPlugins={[remarkGfm]}
              rehypePlugins={[rehypeHighlight]}
            >
              {preview}
            </ReactMarkdown>
          </article>
        </section>
      </main>

      <footer className="app-footer">
        <span>Conçu par</span>
        <a
          href="https://code-wave-eight.vercel.app/"
          target="_blank"
          rel="noreferrer"
        >
          M.G.N CodeWave
        </a>
      </footer>
    </div>
  );
}
