const editor = document.querySelector("#editor");
const statusBadge = document.querySelector("#status");
const fileInput = document.querySelector("#file-input");
const fontSizeSelect = document.querySelector("#font-size");
const pageWidthSelect = document.querySelector("#page-width");

const updateStatus = (message, tone = "success") => {
  statusBadge.textContent = message;
  statusBadge.style.background = tone === "success" ? "#e9f7ef" : "#fff4e5";
  statusBadge.style.color = tone === "success" ? "#0a7b3c" : "#a96500";
};

const exec = (command, value = null) => {
  document.execCommand(command, false, value);
  editor.focus();
};

const applyHeading = (tagName) => {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return;

  const range = selection.getRangeAt(0);
  const container = range.commonAncestorContainer;
  const block = container.nodeType === Node.ELEMENT_NODE
    ? container.closest("p, h1, h2, div")
    : container.parentElement?.closest("p, h1, h2, div");

  if (!block) return;
  const newBlock = document.createElement(tagName);
  newBlock.innerHTML = block.innerHTML;
  block.replaceWith(newBlock);
  editor.focus();
};

const clearDocument = () => {
  editor.innerHTML = "<p></p>";
  updateStatus("New document ready");
};

const handleToolbarClick = (event) => {
  const button = event.target.closest("button");
  if (!button) return;

  const command = button.dataset.command;
  if (command) {
    exec(command);
    return;
  }

  const heading = button.dataset.heading;
  if (heading) {
    applyHeading(heading);
  }
};

const setFontSize = () => {
  exec("fontSize", 7);
  const fontElements = editor.getElementsByTagName("font");
  for (const element of fontElements) {
    if (element.size === "7") {
      element.removeAttribute("size");
      element.style.fontSize = `${fontSizeSelect.value}px`;
    }
  }
};

const updatePageWidth = () => {
  const page = document.querySelector(".page");
  page.style.width = `${pageWidthSelect.value}px`;
};

const triggerOpen = () => fileInput.click();

const createTextRuns = (node, style = {}) => {
  const runs = [];

  if (node.nodeType === Node.TEXT_NODE) {
    if (!node.textContent) return [];
    runs.push(new window.docx.TextRun({
      text: node.textContent,
      bold: style.bold,
      italics: style.italics,
      underline: style.underline ? { type: window.docx.UnderlineType.SINGLE } : undefined,
    }));
    return runs;
  }

  if (node.nodeType !== Node.ELEMENT_NODE) return runs;

  const nextStyle = {
    bold: style.bold || node.tagName === "STRONG" || node.tagName === "B",
    italics: style.italics || node.tagName === "EM" || node.tagName === "I",
    underline: style.underline || node.tagName === "U",
  };

  node.childNodes.forEach((child) => {
    runs.push(...createTextRuns(child, nextStyle));
  });

  return runs;
};

const createParagraphsFromNode = (node) => {
  if (node.nodeType === Node.TEXT_NODE) {
    if (!node.textContent.trim()) return [];
    return [new window.docx.Paragraph({
      children: createTextRuns(node),
    })];
  }

  if (node.nodeType !== Node.ELEMENT_NODE) return [];

  if (node.tagName === "UL") {
    return Array.from(node.querySelectorAll("li")).map((item) =>
      new window.docx.Paragraph({
        children: createTextRuns(item),
        bullet: { level: 0 },
      })
    );
  }

  const headingMap = {
    H1: window.docx.HeadingLevel.TITLE,
    H2: window.docx.HeadingLevel.HEADING_1,
  };

  const heading = headingMap[node.tagName];
  const children = createTextRuns(node);

  if (heading) {
    return [new window.docx.Paragraph({ heading, children })];
  }

  if (node.tagName === "BR") {
    return [new window.docx.Paragraph({})];
  }

  if (["P", "DIV"].includes(node.tagName)) {
    return [new window.docx.Paragraph({ children })];
  }

  const paragraphs = [];
  node.childNodes.forEach((child) => {
    paragraphs.push(...createParagraphsFromNode(child));
  });
  return paragraphs;
};

const exportDocx = async () => {
  updateStatus("Exporting...", "warning");
  const paragraphs = [];
  editor.childNodes.forEach((node) => {
    paragraphs.push(...createParagraphsFromNode(node));
  });

  const doc = new window.docx.Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1000, bottom: 1000, left: 1200, right: 1200 },
          },
        },
        children: paragraphs.length ? paragraphs : [new window.docx.Paragraph({})],
      },
    ],
  });

  const blob = await window.docx.Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "PaperPages.docx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  updateStatus("Exported DOCX");
};

const importDocx = async (file) => {
  updateStatus("Importing...", "warning");
  const arrayBuffer = await file.arrayBuffer();
  const result = await window.mammoth.convertToHtml({ arrayBuffer });
  editor.innerHTML = result.value || "<p></p>";
  updateStatus(`Loaded ${file.name}`);
};

const setupListeners = () => {
  document.querySelector(".toolbar").addEventListener("click", handleToolbarClick);
  document.querySelector("#new-doc").addEventListener("click", clearDocument);
  document.querySelector("#open-doc").addEventListener("click", triggerOpen);
  document.querySelector("#export-doc").addEventListener("click", exportDocx);

  fontSizeSelect.addEventListener("change", setFontSize);
  pageWidthSelect.addEventListener("change", updatePageWidth);

  fileInput.addEventListener("change", (event) => {
    const [file] = event.target.files;
    if (!file) return;
    importDocx(file);
    fileInput.value = "";
  });

  editor.addEventListener("input", () => {
    updateStatus("Editing...");
  });

  editor.addEventListener("blur", () => {
    updateStatus("All changes saved");
  });
};

setupListeners();
