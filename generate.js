const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, LevelFormat, BorderStyle, WidthType,
  ShadingType, ExternalHyperlink, PageBreak
} = require("docx");

const MAROON = "6B1A1A";
const SAFFRON = "C8821A";
const CREAM_DARK = "F0E8D5";
const INK_SOFT = "3D2B1A";
const LIGHT_BG = "FAF5EC";

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "D5C9B5" };
const thinBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

function spacer(pts = 120) {
  return new Paragraph({ spacing: { before: pts, after: 0 }, children: [] });
}

function sectionLabel(text) {
  return new Paragraph({
    spacing: { before: 360, after: 80 },
    children: [new TextRun({ text: text.toUpperCase(), font: "Helvetica Neue", size: 18, bold: true, color: SAFFRON, characterSpacing: 60 })]
  });
}

function bodyText(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 0, after: opts.after || 120 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    children: [new TextRun({ text, font: "Georgia", size: 22, color: INK_SOFT, ...(opts.italic ? { italics: true } : {}), ...(opts.bold ? { bold: true } : {}) })]
  });
}

function richBodyParagraph(runs, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 0, after: opts.after || 120 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    children: runs.map(r => new TextRun({ font: "Georgia", size: 22, color: INK_SOFT, ...r }))
  });
}

function bulletItem(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: Array.isArray(text)
      ? text.map(r => new TextRun({ font: "Georgia", size: 21, color: INK_SOFT, ...r }))
      : [new TextRun({ text, font: "Georgia", size: 21, color: INK_SOFT })]
  });
}

function stageDirection(text) {
  return new Paragraph({
    spacing: { before: 280, after: 80 },
    border: { top: { style: BorderStyle.SINGLE, size: 1, color: SAFFRON, space: 4 } },
    children: [new TextRun({ text: text.toUpperCase(), font: "Helvetica Neue", size: 20, bold: true, color: SAFFRON, characterSpacing: 40 })]
  });
}

function spokenText(text) {
  return new Paragraph({
    spacing: { before: 60, after: 100 },
    children: [new TextRun({ text, font: "Georgia", size: 23, color: INK_SOFT })]
  });
}

function blockQuote(text, cite) {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    indent: { left: 720, right: 720 },
    border: { left: { style: BorderStyle.SINGLE, size: 6, color: SAFFRON, space: 8 } },
    shading: { fill: CREAM_DARK, type: ShadingType.CLEAR },
    children: [
      new TextRun({ text: `\u201C${text}\u201D`, font: "Georgia", size: 26, italics: true, color: MAROON }),
      new TextRun({ text: `  \u2014 ${cite}`, font: "Helvetica Neue", size: 18, color: SAFFRON })
    ]
  });
}

function linkParagraph(label, url, description) {
  return new Paragraph({
    spacing: { before: 80, after: 40 },
    children: [
      new TextRun({ text: `${label}: `, font: "Helvetica Neue", size: 20, bold: true, color: MAROON }),
      new ExternalHyperlink({ link: url, children: [new TextRun({ text: url, font: "Georgia", size: 20, style: "Hyperlink" })] })
    ]
  });
}

function checkboxItem(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: { left: 360 },
    children: Array.isArray(text)
      ? [new TextRun({ text: "\u2610  ", font: "Georgia", size: 24, color: SAFFRON }), ...text.map(r => new TextRun({ font: "Georgia", size: 21, color: INK_SOFT, ...r }))]
      : [new TextRun({ text: "\u2610  ", font: "Georgia", size: 24, color: SAFFRON }), new TextRun({ text, font: "Georgia", size: 21, color: INK_SOFT })]
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Georgia", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: false, font: "Georgia", color: MAROON },
        paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Georgia", color: MAROON },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2014", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { color: SAFFRON, font: "Georgia" } } }]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      // TITLE
      new Paragraph({
        spacing: { before: 200, after: 0 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "\u2767", font: "Georgia", size: 48, color: SAFFRON })]
      }),
      new Paragraph({
        spacing: { before: 200, after: 60 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "37 Practices of a Bodhisattva", font: "Georgia", size: 44, color: MAROON })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "April 13 Prep", font: "Georgia", size: 32, italics: true, color: MAROON })]
      }),
      new Paragraph({
        spacing: { before: 80, after: 200 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "SANGHA STUDY GROUPS  \u00B7  ENGLISH BOOK PUBLICATION", font: "Helvetica Neue", size: 16, color: SAFFRON, characterSpacing: 50 })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 280 },
        alignment: AlignmentType.CENTER,
        border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: SAFFRON, space: 8 } },
        children: []
      }),

      // SECTION 1: STATUS
      sectionLabel("Where We Stand Right Now"),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Project Status \u2014 April 12, 2026")] }),

      // Status table
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 3120, type: WidthType.DXA },
                borders: thinBorders,
                shading: { fill: MAROON, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 160, right: 160 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "$1,100", font: "Georgia", size: 36, color: "E8A84A" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40 }, children: [new TextRun({ text: "PLEDGED TO DATE", font: "Helvetica Neue", size: 14, color: "CCAA88", characterSpacing: 30 })] })
                ]
              }),
              new TableCell({
                width: { size: 3120, type: WidthType.DXA },
                borders: thinBorders,
                shading: { fill: MAROON, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 160, right: 160 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Global", font: "Georgia", size: 36, color: "E8A84A" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40 }, children: [new TextRun({ text: "SINGAPORE GROUP FORMING", font: "Helvetica Neue", size: 14, color: "CCAA88", characterSpacing: 30 })] })
                ]
              }),
              new TableCell({
                width: { size: 3120, type: WidthType.DXA },
                borders: thinBorders,
                shading: { fill: MAROON, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 160, right: 160 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Live", font: "Georgia", size: 36, color: "E8A84A" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40 }, children: [new TextRun({ text: "WEBSITE PUBLISHED", font: "Helvetica Neue", size: 14, color: "CCAA88", characterSpacing: 30 })] })
                ]
              }),
            ]
          })
        ]
      }),

      spacer(200),
      richBodyParagraph([
        { text: "Funding: ", bold: true, color: MAROON },
        { text: "$100 donation received within days of launch. $1,000 pledged by a dharma sister. Funds collected through GBI as fiscal umbrella (501c3, tax-deductible)." }
      ]),
      richBodyParagraph([
        { text: "Groups: ", bold: true, color: MAROON },
        { text: "Singapore sangha reached out, already starting group discussion. US groups forming." }
      ]),
      richBodyParagraph([
        { text: "Translation: ", bold: true, color: MAROON },
        { text: "Begins May 2026. Human translator, paid on completion. Target: end of summer. Estimated cost: $2K\u2013$6K total." }
      ]),

      // SECTION 2: LINKS
      sectionLabel("Live Links"),
      spacer(40),
      linkParagraph("Main Website", "https://bnuckols13.github.io/37-practices/", "Full landing page"),
      bodyText("Full landing page with study plan, facilitator guide, Q&A details, fundraising, and sample invitation.", { indent: 360, after: 160 }),
      linkParagraph("One-Pager / Flyer", "https://bnuckols13.github.io/37-practices/flyer.html", "Shareable"),
      bodyText("Print-friendly overview. Share digitally via this URL or print to PDF.", { indent: 360, after: 160 }),
      linkParagraph("Study Platform", "https://bnuckols13.github.io/37practices/", "Verse-by-verse app"),
      bodyText("Interactive study tool: all 37 verses, transcripts, search, structural frameworks, and analytical toolkit.", { indent: 360, after: 160 }),

      // PAGE BREAK before script
      new Paragraph({ children: [new PageBreak()] }),

      // SECTION 3: SCRIPT
      sectionLabel("Olga\u2019s 4-Minute Presentation Script"),
      spacer(80),

      stageDirection("Opening (30 seconds)"),
      spokenText("Thank you for this opportunity to share something that has become very dear to my heart."),
      spokenText("Many of you know the 37 Practices of a Bodhisattva by Gyalse Tokme Zangpo. Garchen Rinpoche has called it the heart of his teaching and recommends it to all his students."),
      blockQuote("If you follow this text by the letter, you will overcome your suffering.", "Garchen Rinpoche"),
      spokenText("Today I want to tell you about a project that grew from those words."),

      stageDirection("The Project (1 minute)"),
      spokenText("We are building two things together."),
      spokenText("First: peer-led sangha study groups around the world, supported by a structured toolkit so that any community can study this text together. A group in Singapore has already begun forming. Groups are starting in the United States. The toolkit provides session structure, opening practices, facilitator guidance, and a way to ground each gathering in lived experience rather than abstraction."),
      spokenText("Second: the English translation of my commentary on this text. I wrote this book in Ukrainian. I began it before the Russian invasion and completed it during the war\u2019s opening months. It has accompanied readers in bomb shelters, in occupied territories, and in exile. It was recommended alongside Viktor Frankl as one of the most useful mental health resources during the conflict."),

      stageDirection("The Book (1 minute)"),
      spokenText("The commentary includes over 40 original mindfulness exercises, one for each chapter, tested during the most difficult circumstances I have known. It is science-informed. It draws on contemporary research alongside the classical teachings."),
      spokenText("Garchen Rinpoche has personally expressed his wish for this book to reach English readers. Before I left Arizona after our most recent retreat, he presented me with a vajra and bell and told me he wants me to teach."),
      new Paragraph({
        spacing: { before: 120, after: 120 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "\u00B7  \u00B7  \u00B7", font: "Georgia", size: 22, color: SAFFRON })]
      }),
      spokenText("Translation begins in May, with a professional human translator. We expect to complete it by the end of summer."),

      stageDirection("Fundraising (45 seconds)"),
      spokenText("We are raising between $2,000 and $6,000 to cover translation and publication costs. All funds are collected through the Garchen Buddhist Institute, a U.S. nonprofit, to avoid crowdfunding fees and ensure full tax compliance."),
      spokenText("I am grateful to share that we have already received $1,100 in pledges in the first days. Our first donation of $100 arrived almost immediately. A dharma sister has pledged $1,000. The sangha\u2019s generosity is already evident."),

      stageDirection("How to Participate (30 seconds)"),
      spokenText("If you feel moved by this, there are several ways to help."),
      spokenText("You can start or join a study group in your community. You can support the book fund through GBI. You can share this with your sangha network. And once groups are running, I will offer weekly live Q&A sessions, rotating across time zones, so every community has access."),

      stageDirection("Closing (15 seconds)"),
      spokenText("The mission is simple. Not study as an intellectual exercise, but study as refuge. As a direct response to what is actually happening in your life right now."),
      spokenText("Thank you. If you are interested, please speak with me or with Brian and Ryan after this session."),
      spacer(80),
      bodyText("Contact: brianjnuckols@gmail.com", { italic: true }),

      // PAGE BREAK
      new Paragraph({ children: [new PageBreak()] }),

      // SECTION 4: TALKING POINTS
      sectionLabel("Key Talking Points"),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("If someone asks about the book")] }),
      bulletItem("Written in Ukrainian, begun before the invasion, completed during the war\u2019s opening months"),
      bulletItem("Recommended alongside Viktor Frankl as a mental health resource during the conflict"),
      bulletItem("Over 40 original mindfulness exercises, one per chapter"),
      bulletItem("Science-informed: draws on contemporary research alongside classical teachings"),
      bulletItem("Rinpoche personally asked for the English translation. Gave her a vajra and bell. Told her he wants her to teach"),
      bulletItem("Human translator (not machine), professional quality"),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("If someone asks about the study groups")] }),
      bulletItem("10-week format, 90 minutes per session, covering 3\u20134 verses per week"),
      bulletItem([{ text: "Every session opens with: " }, { text: "\u201CWhat do I need refuge from right now in my life?\u201D", italics: true, color: MAROON }]),
      bulletItem("No experience required. 4\u20138 people is ideal. Even two and a text is enough"),
      bulletItem("Toolkit provides full structure: prayers, session plan, facilitator guidance"),
      bulletItem("Recommended commentaries: Dalai Lama, Garchen Rinpoche oral teachings, Geshe Sonam Rinchen, the Karmapa"),
      bulletItem("Weekly Q&A sessions with Olga, rotating across US time zones"),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("If someone asks about money")] }),
      bulletItem("$2K\u2013$6K total estimated cost (editing, cover design, ISBN, print-on-demand)"),
      bulletItem("All funds through GBI (501c3), tax-deductible, no crowdfunding fees"),
      bulletItem("$1,100 already pledged in the first days"),
      bulletItem("Contact: brianjnuckols@gmail.com to pledge or donate"),

      spacer(200),

      // SECTION 5: CHECKLIST
      sectionLabel("What to Send Olga Tonight"),
      spacer(80),
      checkboxItem("The script (copy from Section 3 above, or share this doc)"),
      checkboxItem([{ text: "The flyer link: " }, { text: "bnuckols13.github.io/37-practices/flyer.html", bold: true, color: MAROON }]),
      checkboxItem([{ text: "The main site link: " }, { text: "bnuckols13.github.io/37-practices/", bold: true, color: MAROON }]),
      checkboxItem([{ text: "A short message: " }, { text: "\u201CHere\u2019s your 4-minute script for tomorrow and a one-pager you can share or print. The website is live. You\u2019ve got this.\u201D", italics: true }]),

      spacer(360),
      new Paragraph({
        spacing: { before: 0, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 1, color: SAFFRON, space: 8 } },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Prepared April 12, 2026. All links are live.", font: "Georgia", size: 20, italics: true, color: MAROON })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/Users/briannuckols/Desktop/olga-prep/37-Practices-April-13-Prep.docx", buffer);
  console.log("Document created successfully.");
});
