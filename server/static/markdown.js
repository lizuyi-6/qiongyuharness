function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[char]));
}

function safeUrl(raw) {
  const value = String(raw || "").trim();
  if (!/^https?:\/\/[^\s<>"']+$/i.test(value)) return "";
  return escapeHtml(value);
}

function renderInline(text) {
  const codes = [];
  let source = String(text ?? "").replace(/`([^`\n]+)`/g, (_match, code) => {
    const index = codes.push(`<code>${escapeHtml(code)}</code>`) - 1;
    return `\u0000CODE${index}\u0000`;
  });

  const links = [];
  source = source.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (full, label, href) => {
    const safe = safeUrl(href);
    if (!safe) return full;
    const index = links.push(
      `<a href="${safe}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)}</a>`
    ) - 1;
    return `\u0000LINK${index}\u0000`;
  });

  source = escapeHtml(source);
  source = source.replace(/\*\*([^*\n]+(?:\*(?!\*)[^*\n]*)*)\*\*/g, "<strong>$1</strong>");
  source = source.replace(/~~([^~\n]+)~~/g, "<del>$1</del>");
  source = source.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, "$1<em>$2</em>");

  source = source.replace(/\u0000LINK(\d+)\u0000/g, (_m, i) => links[Number(i)] || "");
  source = source.replace(/\u0000CODE(\d+)\u0000/g, (_m, i) => codes[Number(i)] || "");
  return source;
}

function splitTableRow(line) {
  const trimmed = line.trim().replace(/^\|/, "").replace(/\|$/, "");
  return trimmed.split("|").map((cell) => cell.trim());
}

function isTableSeparator(line) {
  const trimmed = line.trim();
  if (!trimmed.includes("|")) return false;
  const cells = splitTableRow(trimmed);
  return cells.length >= 2 && cells.every((cell) => /^:?-{3,}:?$/.test(cell));
}

function isBlockStart(line, nextLine) {
  const trimmed = line.trim();
  if (!trimmed) return true;
  if (/^```/.test(trimmed)) return true;
  if (/^#{1,6}\s+/.test(trimmed)) return true;
  if (/^(-{3,}|\*{3,}|_{3,})$/.test(trimmed)) return true;
  if (/^>\s?/.test(trimmed)) return true;
  if (/^[-*+]\s+/.test(trimmed)) return true;
  if (/^\d+\.\s+/.test(trimmed)) return true;
  if (trimmed.includes("|") && typeof nextLine === "string" && isTableSeparator(nextLine)) return true;
  return false;
}

export function renderMarkdown(input) {
  const lines = String(input ?? "").replace(/\r\n?/g, "\n").split("\n");
  const blocks = [];
  let index = 0;

  while (index < lines.length) {
    const line = lines[index];
    const trimmed = line.trim();

    if (!trimmed) {
      index += 1;
      continue;
    }

    const fence = trimmed.match(/^```([a-zA-Z0-9_+-]{0,32})?/);
    if (fence) {
      const lang = fence[1] ? ` data-lang="${escapeHtml(fence[1])}"` : "";
      const codeLines = [];
      index += 1;
      while (index < lines.length && !/^```/.test(lines[index].trim())) {
        codeLines.push(lines[index]);
        index += 1;
      }
      if (index < lines.length) index += 1;
      blocks.push(`<pre class="md-code"${lang}><code>${escapeHtml(codeLines.join("\n"))}</code></pre>`);
      continue;
    }

    const heading = trimmed.match(/^(#{1,6})\s+(.+?)\s*#*$/);
    if (heading) {
      const level = Math.min(6, Math.max(3, heading[1].length + 2));
      blocks.push(`<h${level} class="md-heading">${renderInline(heading[2])}</h${level}>`);
      index += 1;
      continue;
    }

    if (/^(-{3,}|\*{3,}|_{3,})$/.test(trimmed)) {
      blocks.push("<hr>");
      index += 1;
      continue;
    }

    if (trimmed.includes("|") && index + 1 < lines.length && isTableSeparator(lines[index + 1])) {
      const headers = splitTableRow(line);
      index += 2;
      const rows = [];
      while (index < lines.length && lines[index].trim().includes("|")) {
        rows.push(splitTableRow(lines[index]));
        index += 1;
      }
      const headHtml = headers.map((cell) => `<th>${renderInline(cell)}</th>`).join("");
      const bodyHtml = rows.map((cells) => {
        const padded = headers.map((_, col) => `<td>${renderInline(cells[col] ?? "")}</td>`).join("");
        return `<tr>${padded}</tr>`;
      }).join("");
      blocks.push(`<div class="table-wrap"><table class="md-table"><thead><tr>${headHtml}</tr></thead><tbody>${bodyHtml}</tbody></table></div>`);
      continue;
    }

    if (/^>\s?/.test(trimmed)) {
      const quoteLines = [];
      while (index < lines.length && /^>\s?/.test(lines[index].trim())) {
        quoteLines.push(lines[index].trim().replace(/^>\s?/, ""));
        index += 1;
      }
      blocks.push(`<blockquote>${quoteLines.map(renderInline).join("<br>")}</blockquote>`);
      continue;
    }

    if (/^[-*+]\s+/.test(trimmed)) {
      const items = [];
      while (index < lines.length && /^[-*+]\s+/.test(lines[index].trim())) {
        items.push(`<li>${renderInline(lines[index].trim().replace(/^[-*+]\s+/, ""))}</li>`);
        index += 1;
      }
      blocks.push(`<ul>${items.join("")}</ul>`);
      continue;
    }

    if (/^\d+\.\s+/.test(trimmed)) {
      const items = [];
      while (index < lines.length && /^\d+\.\s+/.test(lines[index].trim())) {
        items.push(`<li>${renderInline(lines[index].trim().replace(/^\d+\.\s+/, ""))}</li>`);
        index += 1;
      }
      blocks.push(`<ol>${items.join("")}</ol>`);
      continue;
    }

    const paragraph = [];
    while (index < lines.length && !isBlockStart(lines[index], lines[index + 1])) {
      paragraph.push(renderInline(lines[index]));
      index += 1;
    }
    if (paragraph.length) {
      blocks.push(`<p>${paragraph.join("<br>")}</p>`);
    } else {
      blocks.push(`<p>${renderInline(line)}</p>`);
      index += 1;
    }
  }

  return blocks.join("");
}
