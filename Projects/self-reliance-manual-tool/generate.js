/**
 * generate.js
 * Self-Reliance Manual Content Entry Tool
 *
 * Uses the docx library (loaded via CDN as window.docx)
 * to produce a formatted .docx file from the HTML form.
 *
 * Template structure per lesson:
 *   EVALUATE → LEARN → PRACTICE → SERVE
 */

/* ── Constants ── */
const FONT = "Open Sans";
const PAGE_W = 12240;      // 8.5 in in DXA
const PAGE_H = 15840;      // 11 in in DXA
const MARGIN = 1080;       // 0.75 in margins
const CONTENT_W = PAGE_W - MARGIN * 2; // 10080 DXA

// Section accent colors (hex, no #)
const COLOR = {
  evaluate: "1A6EA8",
  learn:    "2E7D32",
  practice: "B45309",
  serve:    "7B3FA0",
  gray:     "F3F4F6",
  grayText: "4B5563",
  muted:    "6B7280",
  black:    "1A1A2E",
  white:    "FFFFFF",
  border:   "D1D5DB",
  headerBg: "1A1A2E",
};

/* ── Helpers ── */

/**
 * Returns a docx TextRun with the given text and options.
 */
function run(text, opts = {}) {
  const { docx } = window;
  return new docx.TextRun({
    text: String(text || ""),
    font: FONT,
    size: opts.size || 22,           // pt × 2 (half-points); 22 = 11pt default
    bold: opts.bold || false,
    italics: opts.italics || false,
    color: opts.color || COLOR.black,
    ...opts._extra,
  });
}

/**
 * A simple paragraph builder.
 */
function para(children, opts = {}) {
  const { docx } = window;
  const runs = Array.isArray(children) ? children : [children];
  return new docx.Paragraph({
    children: runs,
    spacing: {
      before: opts.before !== undefined ? opts.before : 80,
      after:  opts.after  !== undefined ? opts.after  : 80,
      line: opts.line || 360,         // 1.5 line spacing (240 = single)
    },
    alignment: opts.align || docx.AlignmentType.LEFT,
    ...(opts._extra || {}),
  });
}

/** Empty spacer paragraph */
function spacer(pts = 4) {
  return para([run("")], { before: 0, after: pts * 20 });
}

/** Bold section heading (EVALUATE / LEARN / PRACTICE / SERVE) */
function sectionHeading(text, color) {
  const { docx } = window;
  return new docx.Paragraph({
    children: [
      new docx.TextRun({
        text,
        font: FONT,
        size: 52,          // 26pt
        bold: true,
        color: COLOR.white,
        _extra: {},
      }),
    ],
    shading: { fill: color, type: docx.ShadingType.CLEAR },
    spacing: { before: 200, after: 120, line: 360 },
    indent: { left: 160, right: 160 },
  });
}

/** Sub-time note "(8 min)" appended to section heading */
function sectionHeadingWithTime(label, time, color) {
  const { docx } = window;
  return new docx.Paragraph({
    children: [
      new docx.TextRun({ text: label, font: FONT, size: 52, bold: true, color: COLOR.white }),
      new docx.TextRun({ text: `  (${time})`, font: FONT, size: 32, bold: false, color: "D1D5DB" }),
    ],
    shading: { fill: color, type: docx.ShadingType.CLEAR },
    spacing: { before: 200, after: 120, line: 360 },
    indent: { left: 160, right: 160 },
  });
}

/** "Maximum Time" sub-line beneath section heading */
function maxTimePara(text) {
  return para([
    run(text, { size: 32, bold: false, color: COLOR.muted, italics: true }),
  ], { before: 0, after: 60 });
}

/** Bold step label e.g. "STEP 1: EVALUATE YOUR EFFORTS  (5 minutes)" */
function stepLabel(stepNum, title, timeNote = "") {
  const { docx } = window;
  const children = [
    new docx.TextRun({ text: `STEP ${stepNum}: `, font: FONT, size: 28, bold: true, color: COLOR.black }),
    new docx.TextRun({ text: title, font: FONT, size: 28, bold: true, color: COLOR.black }),
  ];
  if (timeNote) {
    children.push(new docx.TextRun({ text: `  (${timeNote})`, font: FONT, size: 24, bold: false, color: COLOR.muted }));
  }
  return para(children, { before: 160, after: 80 });
}

/** Italic "Read:" label */
function readLabel() {
  const { docx } = window;
  return new docx.Paragraph({
    children: [new docx.TextRun({ text: "Read:", font: FONT, size: 28, bold: true, color: COLOR.black })],
    spacing: { before: 140, after: 60, line: 360 },
  });
}

/** "Discuss:" label */
function discussLabel() {
  const { docx } = window;
  return new docx.Paragraph({
    children: [new docx.TextRun({ text: "Discuss:", font: FONT, size: 28, bold: true, color: COLOR.black })],
    spacing: { before: 140, after: 60, line: 360 },
  });
}

/** "Watch:" label */
function watchLabel() {
  const { docx } = window;
  return new docx.Paragraph({
    children: [new docx.TextRun({ text: "Watch:", font: FONT, size: 28, bold: true, color: COLOR.black })],
    spacing: { before: 140, after: 60, line: 360 },
  });
}

/** "Ponder:" label */
function ponderLabel() {
  const { docx } = window;
  return new docx.Paragraph({
    children: [new docx.TextRun({ text: "Ponder:", font: FONT, size: 28, bold: true, color: COLOR.black })],
    spacing: { before: 140, after: 60, line: 360 },
  });
}

/** Body text paragraph at 17pt */
function bodyPara(text, opts = {}) {
  return para([run(text, { size: 34, ...opts })], { before: 60, after: 60, line: 360 });
}

/** Scripture / quote block — italicized, gray background */
function quoteBlock(quoteText, source) {
  const { docx } = window;
  const children = [];

  if (quoteText && quoteText.trim()) {
    children.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: quoteText, font: FONT, size: 36, italics: true, color: COLOR.black })],
        shading: { fill: COLOR.gray, type: docx.ShadingType.CLEAR },
        spacing: { before: 120, after: 60, line: 360 },
        indent: { left: 240, right: 240 },
      })
    );
  }

  if (source && source.trim()) {
    children.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: source, font: FONT, size: 32, italics: false, color: COLOR.muted })],
        shading: { fill: COLOR.gray, type: docx.ShadingType.CLEAR },
        spacing: { before: 0, after: 120, line: 360 },
        indent: { left: 360, right: 240 },
      })
    );
  }

  return children;
}

/** Numbered section heading inside LEARN */
function learnSectionTitle(num, text) {
  const { docx } = window;
  return new docx.Paragraph({
    children: [
      new docx.TextRun({ text: `${num}. `, font: FONT, size: 40, bold: true, color: COLOR.black }),
      new docx.TextRun({ text: text.toUpperCase(), font: FONT, size: 40, bold: true, color: COLOR.black }),
    ],
    spacing: { before: 240, after: 100, line: 360 },
  });
}

/** "TODAY WE WILL LEARN" box */
function todayWillLearnBox(principle1, principle2) {
  const { docx } = window;
  const border = { style: docx.BorderStyle.SINGLE, size: 4, color: COLOR.border };
  const borders = { top: border, bottom: border, left: border, right: border };

  const cellChildren = [
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "TODAY WE WILL LEARN", font: FONT, size: 40, bold: true, color: COLOR.black })],
      spacing: { before: 80, after: 60, line: 360 },
    }),
  ];

  if (principle1 && principle1.trim()) {
    cellChildren.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: principle1, font: FONT, size: 36, bold: true, color: COLOR.black })],
        spacing: { before: 40, after: 40, line: 360 },
      })
    );
  }
  if (principle2 && principle2.trim()) {
    cellChildren.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: principle2, font: FONT, size: 36, bold: false, color: COLOR.black })],
        spacing: { before: 0, after: 80, line: 360 },
      })
    );
  }

  return new docx.Table({
    width: { size: CONTENT_W, type: docx.WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new docx.TableRow({
        children: [
          new docx.TableCell({
            borders,
            width: { size: CONTENT_W, type: docx.WidthType.DXA },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            children: cellChildren,
          }),
        ],
      }),
    ],
  });
}

/** Evaluate effort table */
function evaluateTable(commitments) {
  const { docx } = window;
  const headerBorder = { style: docx.BorderStyle.SINGLE, size: 4, color: COLOR.border };
  const borders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

  // Header row
  const headerRow = new docx.TableRow({
    tableHeader: true,
    children: [
      makeHeaderCell('THIS LAST WEEK I PRACTICED: Indicate your effort level (Choose \u201c1\u201d for little, \u201c2\u201d for some, or \u201c3\u201d for significant)', 6400, COLOR.headerBg),
      makeHeaderCell("Effort Level", 1600, COLOR.headerBg),
      makeHeaderCell("Action Partner Initials", 2080, COLOR.headerBg),
    ],
  });

  const rows = [headerRow];
  const labels = ["A", "B", "C", "D"];
  const defaultTexts = [
    commitments[0] || "(restate practice commitments from prior week)",
    commitments[1] || "Objective 2",
    commitments[2] || "Objective 3",
    "Custom objective",
  ];

  labels.forEach((label, i) => {
    rows.push(
      new docx.TableRow({
        children: [
          new docx.TableCell({
            borders,
            width: { size: 6400, type: docx.WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({ text: label + "  ", font: FONT, size: 28, bold: true }),
                  new docx.TextRun({ text: defaultTexts[i], font: FONT, size: 28 }),
                ],
                spacing: { before: 40, after: 40, line: 320 },
              }),
            ],
          }),
          new docx.TableCell({
            borders,
            width: { size: 1600, type: docx.WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              new docx.Paragraph({
                children: [new docx.TextRun({ text: "1    2    3", font: FONT, size: 28, color: COLOR.muted })],
                spacing: { before: 40, after: 40, line: 320 },
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            borders,
            width: { size: 2080, type: docx.WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new docx.Paragraph({ children: [new docx.TextRun("")], spacing: { before: 40, after: 40 } })],
          }),
        ],
      })
    );
  });

  return new docx.Table({
    width: { size: CONTENT_W, type: docx.WidthType.DXA },
    columnWidths: [6400, 1600, 2080],
    rows,
  });
}

function makeHeaderCell(text, width, bgColor) {
  const { docx } = window;
  const border = { style: docx.BorderStyle.SINGLE, size: 4, color: COLOR.border };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new docx.TableCell({
    borders,
    width: { size: width, type: docx.WidthType.DXA },
    shading: { fill: bgColor, type: docx.ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [
      new docx.Paragraph({
        children: [new docx.TextRun({ text, font: FONT, size: 26, bold: true, color: COLOR.white })],
        spacing: { before: 40, after: 40, line: 300 },
      }),
    ],
  });
}

/** Action partner name table */
function actionPartnerTable() {
  const { docx } = window;
  const border = { style: docx.BorderStyle.SINGLE, size: 4, color: COLOR.border };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new docx.Table({
    width: { size: CONTENT_W, type: docx.WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new docx.TableRow({
        children: [
          new docx.TableCell({
            borders,
            width: { size: CONTENT_W, type: docx.WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 200, right: 200 },
            children: [
              new docx.Paragraph({
                children: [new docx.TextRun({ text: "Action Partner's Name:", font: FONT, size: 28, bold: true })],
                spacing: { before: 40, after: 160, line: 320 },
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

/** Practice commitment table */
function practiceTable(commitments) {
  const { docx } = window;
  const border = { style: docx.BorderStyle.SINGLE, size: 4, color: COLOR.border };
  const borders = { top: border, bottom: border, left: border, right: border };
  const colMain = CONTENT_W - 2000;
  const colWhen = 2000;

  const headerRow = new docx.TableRow({
    tableHeader: true,
    children: [
      makeHeaderCell("I WILL PRACTICE BY:", colMain, COLOR.headerBg),
      makeHeaderCell("BY WHEN", colWhen, COLOR.headerBg),
    ],
  });

  const labels = ["A", "B", "C", "D"];
  const texts = [
    commitments[0] || "Practice commitment A",
    commitments[1] || "Practice commitment B",
    commitments[2] || "Practice commitment C",
    "(Leave blank — write your own inspired idea here)",
  ];

  const dataRows = labels.map((label, i) =>
    new docx.TableRow({
      children: [
        new docx.TableCell({
          borders,
          width: { size: colMain, type: docx.WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: label + "  ", font: FONT, size: 28, bold: true }),
                new docx.TextRun({ text: texts[i], font: FONT, size: 28, italics: i === 3, color: i === 3 ? COLOR.muted : COLOR.black }),
              ],
              spacing: { before: 40, after: 40, line: 320 },
            }),
          ],
        }),
        new docx.TableCell({
          borders,
          width: { size: colWhen, type: docx.WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new docx.Paragraph({ children: [new docx.TextRun("")], spacing: { before: 40, after: 160 } })],
        }),
      ],
    })
  );

  return new docx.Table({
    width: { size: CONTENT_W, type: docx.WidthType.DXA },
    columnWidths: [colMain, colWhen],
    rows: [headerRow, ...dataRows],
  });
}

/** Serve commitment table */
function serveTable(customExample) {
  const { docx } = window;
  const border = { style: docx.BorderStyle.SINGLE, size: 4, color: COLOR.border };
  const borders = { top: border, bottom: border, left: border, right: border };
  const colMain = CONTENT_W - 2000;
  const colWhen = 2000;

  const headerRow = new docx.TableRow({
    tableHeader: true,
    children: [
      makeHeaderCell("I WILL SERVE BY:", colMain, COLOR.headerBg),
      makeHeaderCell("BY WHEN", colWhen, COLOR.headerBg),
    ],
  });

  const rows = [
    ["A", "Contacting my action partner this week to support them in accomplishing their commitments.", false],
    ["B", customExample ? `Other: ${customExample}` : "Other:", false],
  ];

  const dataRows = rows.map(([label, text]) =>
    new docx.TableRow({
      children: [
        new docx.TableCell({
          borders,
          width: { size: colMain, type: docx.WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: label + "  ", font: FONT, size: 28, bold: true }),
                new docx.TextRun({ text, font: FONT, size: 28 }),
              ],
              spacing: { before: 40, after: 40, line: 320 },
            }),
          ],
        }),
        new docx.TableCell({
          borders,
          width: { size: colWhen, type: docx.WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new docx.Paragraph({ children: [new docx.TextRun("")], spacing: { before: 40, after: 160 } })],
        }),
      ],
    })
  );

  return new docx.Table({
    width: { size: CONTENT_W, type: docx.WidthType.DXA },
    columnWidths: [colMain, colWhen],
    rows: [headerRow, ...dataRows],
  });
}

/** Signature table */
function signatureTable() {
  const { docx } = window;
  const border = { style: docx.BorderStyle.SINGLE, size: 4, color: COLOR.border };
  const borders = { top: border, bottom: border, left: border, right: border };
  const colW = Math.floor(CONTENT_W / 3);
  const colW3 = CONTENT_W - colW * 2;

  const makeSignCell = (label, w) =>
    new docx.TableCell({
      borders,
      width: { size: w, type: docx.WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [
        new docx.Paragraph({
          children: [new docx.TextRun({ text: label + ":", font: FONT, size: 26, bold: true })],
          spacing: { before: 40, after: 200, line: 320 },
        }),
      ],
    });

  return new docx.Table({
    width: { size: CONTENT_W, type: docx.WidthType.DXA },
    columnWidths: [colW, colW, colW3],
    rows: [
      new docx.TableRow({
        children: [
          makeSignCell("My signature", colW),
          makeSignCell("Action partner's signature", colW),
          makeSignCell("Contact information", colW3),
        ],
      }),
    ],
  });
}

/** Generic 3-column activity table from form data */
function activityTable(headers, rows) {
  const { docx } = window;

  // Skip if no meaningful content
  const hasContent = headers.some(h => h.trim()) || rows.some(r => r.some(c => c.trim()));
  if (!hasContent) return null;

  const border = { style: docx.BorderStyle.SINGLE, size: 4, color: COLOR.border };
  const borders = { top: border, bottom: border, left: border, right: border };
  const colW = Math.floor(CONTENT_W / 3);
  const colW3 = CONTENT_W - colW * 2;
  const colWidths = [colW, colW, colW3];

  const makeCell = (text, w, isHeader = false) =>
    new docx.TableCell({
      borders,
      width: { size: w, type: docx.WidthType.DXA },
      shading: isHeader ? { fill: COLOR.gray, type: docx.ShadingType.CLEAR } : undefined,
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [
        new docx.Paragraph({
          children: [new docx.TextRun({ text: text || "", font: FONT, size: 28, bold: isHeader })],
          spacing: { before: 40, after: 40, line: 320 },
        }),
      ],
    });

  const tableRows = [];

  if (headers.some(h => h.trim())) {
    tableRows.push(
      new docx.TableRow({
        tableHeader: true,
        children: headers.map((h, i) => makeCell(h, colWidths[i], true)),
      })
    );
  }

  rows.forEach(row => {
    if (row.some(c => c.trim())) {
      tableRows.push(
        new docx.TableRow({
          children: row.map((c, i) => makeCell(c, colWidths[i])),
        })
      );
    }
  });

  if (tableRows.length === 0) return null;

  return new docx.Table({
    width: { size: CONTENT_W, type: docx.WidthType.DXA },
    columnWidths: [colW, colW, colW3],
    rows: tableRows,
  });
}

/** Bullet list item */
function bulletItem(text) {
  const { docx } = window;
  return new docx.Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new docx.TextRun({ text, font: FONT, size: 34 })],
    spacing: { before: 60, after: 60, line: 360 },
  });
}

/** Page break */
function pageBreak() {
  const { docx } = window;
  return new docx.Paragraph({ children: [new docx.PageBreak()] });
}

/* ── Collect form data ── */
function collectFormData() {
  const v = id => (document.getElementById(id)?.value || "").trim();
  const all = cls => [...document.querySelectorAll(`.${cls}`)].map(el => el.value.trim());

  // Activity table
  const tableColHeaders = [...document.querySelectorAll(".table-col-header")].map(el => el.value.trim());
  const tableDataRows = [...document.querySelectorAll("#activity-table-rows .table-data-row")].map(row =>
    [...row.querySelectorAll("input")].map(inp => inp.value.trim())
  );

  return {
    lessonNumber: v("lessonNumber"),
    lessonTitle: v("lessonTitle"),
    psh: all("psh-item").filter(Boolean),

    // Evaluate
    evalCommitments: [v("evalA"), v("evalB"), v("evalC")],
    evaluateDiscuss: v("evaluateDiscuss"),

    // Learn
    learnPrinciple1: v("learnPrinciple1"),
    learnPrinciple2: v("learnPrinciple2"),

    learnSection1Title: v("learnSection1Title"),
    learnSection1Read1: v("learnSection1Read1"),
    learnSection1Quote: v("learnSection1Quote"),
    learnSection1QuoteSource: v("learnSection1QuoteSource"),
    learnSection1Discuss: v("learnSection1Discuss"),
    learnSection1Read2: v("learnSection1Read2"),

    learnSection2Title: v("learnSection2Title"),
    learnSection2Read1: v("learnSection2Read1"),
    learnSection2Quote: v("learnSection2Quote"),
    learnSection2QuoteSource: v("learnSection2QuoteSource"),
    learnSection2Discuss: v("learnSection2Discuss"),

    videoTitle: v("videoTitle"),
    videoUrl: v("videoUrl"),
    videoTranscript: v("videoTranscript"),

    activityStep1: v("activityStep1"),
    activityStep2: v("activityStep2"),
    activityStep3: v("activityStep3"),
    activityTableHeaders: tableColHeaders,
    activityTableRows: tableDataRows,
    activityPonder: v("activityPonder"),

    // Practice
    practiceCommitments: [v("practiceA"), v("practiceB"), v("practiceC")],

    // Serve
    serveCustomExample: v("serveCustomExample"),
  };
}

/* ── Build the document ── */
function buildDocument(d) {
  const { docx } = window;

  const children = [];

  /* ===== LESSON TITLE ===== */
  children.push(
    new docx.Paragraph({
      children: [
        new docx.TextRun({ text: `LESSON ${d.lessonNumber}`, font: FONT, size: 48, bold: true, color: COLOR.muted }),
      ],
      spacing: { before: 0, after: 60, line: 360 },
    }),
    new docx.Paragraph({
      children: [new docx.TextRun({ text: d.lessonTitle.toUpperCase() || "LESSON TITLE", font: FONT, size: 80, bold: true, color: COLOR.black })],
      spacing: { before: 0, after: 120, line: 400 },
    })
  );

  /* Principles, Skills, and Habits */
  if (d.psh.length > 0) {
    children.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "PRINCIPLES, SKILLS, AND HABITS", font: FONT, size: 34, bold: true, color: COLOR.muted })],
        spacing: { before: 120, after: 80, line: 360 },
      })
    );
    d.psh.forEach(item => children.push(bulletItem(item)));
  }

  children.push(spacer(12));

  /* ===== EVALUATE ===== */
  children.push(sectionHeadingWithTime("EVALUATE", "8 minutes", COLOR.evaluate));
  children.push(maxTimePara("Maximum Time: 8 minutes"));

  children.push(stepLabel("1", "EVALUATE YOUR EFFORTS", "5 minutes"));
  children.push(bodyPara("Take a minute to think about how you did with your commitments last week."));
  children.push(spacer(4));
  children.push(evaluateTable(d.evalCommitments));
  children.push(spacer(8));

  children.push(stepLabel("2", "SHARE YOUR EXPERIENCE", "5 minutes"));
  children.push(bodyPara("As a group, share the things you learned last week while working on your commitments. What successes did you experience? What difficulties? The questions below might help the discussion. Going forward, commit to continue to practice these skills."));

  if (d.evaluateDiscuss) {
    children.push(discussLabel());
    children.push(bodyPara(d.evaluateDiscuss));
  }

  children.push(spacer(8));

  children.push(stepLabel("3", "CHOOSE AN ACTION PARTNER", "2 minutes"));
  children.push(bodyPara("Choose an action partner from the group for this week. Generally, action partners are the same gender and are not family members. Take a couple of minutes now to meet with your action partner. Introduce yourselves and discuss how you will contact each other through the week."));
  children.push(spacer(4));
  children.push(actionPartnerTable());
  children.push(spacer(12));

  /* ===== LEARN ===== */
  children.push(sectionHeadingWithTime("LEARN", "40 minutes", COLOR.learn));
  children.push(maxTimePara("Maximum Time: 40 minutes"));
  children.push(spacer(6));
  children.push(todayWillLearnBox(d.learnPrinciple1, d.learnPrinciple2));
  children.push(spacer(10));

  /* Learn Section 1 */
  if (d.learnSection1Title) {
    children.push(learnSectionTitle(1, d.learnSection1Title));
  }

  if (d.learnSection1Read1) {
    children.push(readLabel());
    children.push(bodyPara(d.learnSection1Read1));
  }

  if (d.learnSection1Quote) {
    children.push(...quoteBlock(d.learnSection1Quote, d.learnSection1QuoteSource));
  }

  if (d.learnSection1Discuss) {
    children.push(discussLabel());
    children.push(bodyPara(d.learnSection1Discuss));
  }

  if (d.learnSection1Read2) {
    children.push(readLabel());
    children.push(bodyPara(d.learnSection1Read2));
  }

  children.push(spacer(8));

  /* Learn Section 2 */
  if (d.learnSection2Title) {
    children.push(learnSectionTitle(2, d.learnSection2Title));
  }

  if (d.learnSection2Read1) {
    children.push(readLabel());
    children.push(bodyPara(d.learnSection2Read1));
  }

  if (d.learnSection2Quote) {
    children.push(...quoteBlock(d.learnSection2Quote, d.learnSection2QuoteSource));
  }

  if (d.learnSection2Discuss) {
    children.push(discussLabel());
    children.push(bodyPara(d.learnSection2Discuss));
  }

  children.push(spacer(8));

  /* Watch / Video */
  if (d.videoTitle || d.videoUrl || d.videoTranscript) {
    children.push(watchLabel());

    if (d.videoTitle && d.videoUrl) {
      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: 'Watch the video "', font: FONT, size: 34 }),
            new docx.ExternalHyperlink({
              link: d.videoUrl,
              children: [new docx.TextRun({ text: d.videoTitle, font: FONT, size: 34, style: "Hyperlink" })],
            }),
            new docx.TextRun({ text: '." ', font: FONT, size: 34 }),
            new docx.TextRun({ text: "(Available at ChurchofJesusChrist.org/media. If you are unable to watch the video, you can access the transcript below.)", font: FONT, size: 30, italics: true, color: COLOR.muted }),
          ],
          spacing: { before: 60, after: 60, line: 360 },
        })
      );
    } else if (d.videoTitle) {
      children.push(bodyPara(`Watch the video "${d.videoTitle}."`));
    }

    if (d.videoTranscript) {
      children.push(spacer(6));
      children.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: `${d.videoTitle ? d.videoTitle.toUpperCase() : "VIDEO"} (transcript)`, font: FONT, size: 40, bold: true, color: COLOR.black })],
          spacing: { before: 140, after: 80, line: 360 },
        })
      );
      children.push(bodyPara(d.videoTranscript));
    }

    children.push(spacer(8));
  }

  /* Activity */
  const hasActivity = d.activityStep1 || d.activityStep2 || d.activityStep3 || d.activityPonder;
  if (hasActivity || d.activityTableHeaders.some(h => h)) {
    children.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "ACTIVITY ", font: FONT, size: 34, bold: true }), new docx.TextRun({ text: "(4 minutes)", font: FONT, size: 28, color: COLOR.muted })],
        spacing: { before: 140, after: 80, line: 360 },
      })
    );

    if (d.activityStep1) {
      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: "Step 1: ", font: FONT, size: 32, bold: true }),
            new docx.TextRun({ text: d.activityStep1, font: FONT, size: 32 }),
          ],
          spacing: { before: 80, after: 60, line: 360 },
        })
      );
    }

    if (d.activityStep2) {
      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: "Step 2: ", font: FONT, size: 32, bold: true }),
            new docx.TextRun({ text: d.activityStep2, font: FONT, size: 32 }),
          ],
          spacing: { before: 80, after: 60, line: 360 },
        })
      );
    }

    const tbl = activityTable(d.activityTableHeaders, d.activityTableRows);
    if (tbl) {
      children.push(spacer(4));
      children.push(tbl);
      children.push(spacer(4));
    }

    if (d.activityStep3) {
      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: "Step 3: ", font: FONT, size: 32, bold: true }),
            new docx.TextRun({ text: d.activityStep3, font: FONT, size: 32 }),
          ],
          spacing: { before: 80, after: 60, line: 360 },
        })
      );
    }

    if (d.activityPonder) {
      children.push(ponderLabel());
      children.push(bodyPara(d.activityPonder));
    }

    children.push(spacer(10));
  }

  /* ===== PRACTICE ===== */
  children.push(sectionHeadingWithTime("PRACTICE", "8 minutes", COLOR.practice));
  children.push(maxTimePara("Maximum Time: 8 minutes"));

  children.push(readLabel());
  children.push(bodyPara("Below make commitments to practice what you have learned. Set a date by when you will complete your practice commitment. When the group meets next week, each group member will begin by evaluating their successes and struggles they experienced while practicing."));
  children.push(bodyPara("In the blank space below, write any idea you feel inspired to practice this week. This additional commitment is not required but be open to feelings you may have on areas you need to work on this week. This might include a concept from a prior week or a spiritual goal such as improving daily prayer or scripture study."));

  children.push(spacer(6));
  children.push(practiceTable(d.practiceCommitments));
  children.push(spacer(12));

  /* ===== SERVE ===== */
  children.push(sectionHeadingWithTime("SERVE", "8 minutes", COLOR.serve));
  children.push(maxTimePara("Maximum Time: 8 minutes"));

  children.push(readLabel());
  children.push(bodyPara("Your action partner you chose at the beginning of today's lesson is your accountability partner who will support you in accomplishing your practice commitments this week."));
  children.push(bodyPara("Share your practice commitments with one another. Then complete the serve commitments below. Determine how and when you will follow up and support one another this week."));

  if (d.serveCustomExample) {
    children.push(bodyPara(`Add another serve goal in the blank line. You might think about ${d.serveCustomExample}. Then sign below.`));
  }

  children.push(spacer(6));
  children.push(serveTable(d.serveCustomExample));
  children.push(spacer(8));
  children.push(signatureTable());
  children.push(spacer(12));

  /* Closing */
  children.push(
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "Facilitator: ", font: FONT, size: 28, bold: true }), new docx.TextRun({ text: "Gather the group. Remind everyone of the next date and time the group will meet.", font: FONT, size: 28 })],
      spacing: { before: 80, after: 80, line: 360 },
    }),
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "Closing prayer.", font: FONT, size: 28, italics: true })],
      spacing: { before: 60, after: 60, line: 360 },
    })
  );

  /* ===== ASSEMBLE DOCUMENT ===== */
  return new docx.Document({
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [
            {
              level: 0,
              format: docx.LevelFormat.BULLET,
              text: "\u2022",
              alignment: docx.AlignmentType.LEFT,
              style: {
                paragraph: { indent: { left: 720, hanging: 360 } },
              },
            },
          ],
        },
      ],
    },
    styles: {
      default: {
        document: {
          run: { font: FONT, size: 22 },
        },
      },
      paragraphStyles: [
        {
          id: "Hyperlink",
          name: "Hyperlink",
          basedOn: "Normal",
          run: { color: "1A6EA8", underline: { type: docx.UnderlineType.SINGLE } },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE_W, height: PAGE_H },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        children,
      },
    ],
  });
}

/* ── Generate and download ── */
async function generateDocx() {
  const btn = document.getElementById("generateBtn");
  const status = document.getElementById("status-msg");

  btn.disabled = true;
  status.textContent = "Generating...";
  status.className = "status-msg";

  try {
    const data = collectFormData();

    if (!data.lessonTitle) {
      status.textContent = "Please enter a Lesson Title before generating.";
      status.className = "status-msg error";
      btn.disabled = false;
      return;
    }

    const doc = buildDocument(data);
    const buffer = await window.docx.Packer.toBuffer(doc);
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Lesson-${data.lessonNumber}-${data.lessonTitle.replace(/\s+/g, "-")}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    status.textContent = `✓ Downloaded: Lesson-${data.lessonNumber}-${data.lessonTitle.replace(/\s+/g, "-")}.docx`;
    status.className = "status-msg success";
  } catch (err) {
    console.error("Generation error:", err);
    status.textContent = `Error: ${err.message || "Something went wrong. Check the console."}`;
    status.className = "status-msg error";
  } finally {
    btn.disabled = false;
  }
}

/* ── Wire up ── */
document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("generateBtn").addEventListener("click", generateDocx);

  /* Dynamic "Add Row" button for activity table */
  document.getElementById("add-table-row").addEventListener("click", () => {
    const container = document.getElementById("activity-table-rows");
    const rowCount = container.querySelectorAll(".table-data-row").length + 1;
    const div = document.createElement("div");
    div.className = "table-data-row";
    div.innerHTML = `
      <input type="text" placeholder="Row ${rowCount}, Col 1" />
      <input type="text" placeholder="Row ${rowCount}, Col 2" />
      <input type="text" placeholder="Row ${rowCount}, Col 3" />
    `;
    container.appendChild(div);
  });
});
