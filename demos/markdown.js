import { setup } from './shared.js';
import source from './markdown.js?raw';

// #region demo
import { marked } from 'marked';
import DOMPurify from 'dompurify';
import hljs from 'highlight.js';

async function render(buffer, name, out) {
  const text = new TextDecoder().decode(buffer);

  const article = document.createElement('article');
  article.className = 'doc-html';

  if (name.endsWith('.md')) {
    // Markdown -> HTML, then sanitize: Markdown may contain raw HTML/script.
    article.innerHTML = DOMPurify.sanitize(await marked.parse(text));
    article.querySelectorAll('pre code').forEach((el) => hljs.highlightElement(el));
  } else {
    // Any code/text file: syntax-highlighted, language auto-detected.
    const code = document.createElement('code');
    code.textContent = text;
    article.appendChild(document.createElement('pre')).append(code);
    hljs.highlightElement(code);
  }
  out.append(article);
}
// #endregion

setup({ sample: 'notes.md', accept: '.md,.txt,.json,.xml,.yaml,.yml,.js,.ts,.css,.sql,.html,.log,.ini', render, source });
