const fs = require('fs');
const path = require('path');
const { 
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
  WidthType, AlignmentType, BorderStyle, HeadingLevel, ShadingType 
} = require('docx');

// Colors
const COLOR_PRIMARY = "4F46E5";   // Indigo 600
const COLOR_PRIMARY_DARK = "312E81"; // Indigo 900
const COLOR_DARK = "0F172A";      // Slate 900
const COLOR_MUTED = "475569";     // Slate 600
const COLOR_LIGHT_BG = "F8FAFC";  // Slate 50
const COLOR_BORDER = "CBD5E1";    // Slate 300
const COLOR_ACCENT = "059669";    // Emerald 600

function createHeaderBanner() {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [
      new TextRun({
        text: "TechNest – Smart Electronics Store",
        bold: true,
        size: 26, // 13pt
        color: COLOR_PRIMARY,
        font: "Calibri",
      }),
      new TextRun({
        text: "\nTechNest – E-Commerce Website: Framework and Technologies Used",
        bold: true,
        size: 30, // 15pt
        color: COLOR_PRIMARY_DARK,
        font: "Calibri",
      }),
      new TextRun({
        text: "\nMBA E-Commerce Academic Project Specification | One-Page Technical Brief",
        italics: true,
        size: 18, // 9pt
        color: COLOR_MUTED,
        font: "Calibri",
      }),
    ],
  });
}

function createSectionHeading(title) {
  return new Paragraph({
    spacing: { before: 120, after: 50 },
    border: {
      bottom: {
        color: COLOR_PRIMARY,
        size: 12,
        style: BorderStyle.SINGLE,
      },
    },
    children: [
      new TextRun({
        text: title.toUpperCase(),
        bold: true,
        size: 20, // 10pt
        color: COLOR_PRIMARY_DARK,
        font: "Calibri",
      }),
    ],
  });
}

function createParagraph(text, isBold = false) {
  return new Paragraph({
    spacing: { before: 30, after: 60 },
    alignment: AlignmentType.JUSTIFY,
    children: [
      new TextRun({
        text: text,
        bold: isBold,
        size: 19, // 9.5pt
        color: COLOR_DARK,
        font: "Calibri",
      }),
    ],
  });
}

function createTechTable() {
  const tableData = [
    {
      domain: "Frontend Framework",
      tech: "React.js (v18.3.1)",
      desc: "Declarative, component-based UI architecture with hooks & functional components"
    },
    {
      domain: "Core Web Standards",
      tech: "HTML5, CSS3, JavaScript (ES6+ / JSX)",
      desc: "Modern semantic web standards, native DOM APIs, and dynamic JSX templating"
    },
    {
      domain: "Styling & Design System",
      tech: "Tailwind CSS (v3.4.15), PostCSS, Autoprefixer",
      desc: "Utility-first CSS framework with responsive utility classes and CSS variables"
    },
    {
      domain: "Icons & Visual Elements",
      tech: "Lucide React (v0.460.0), Canvas-Confetti",
      desc: "Feather-derived SVG icon set and canvas celebratory checkout animation"
    },
    {
      domain: "Build Tool & Bundler",
      tech: "Vite (v5.4.10) with @vitejs/plugin-react",
      desc: "Next-generation ESM native frontend dev server with Rollup production bundler"
    },
    {
      domain: "State & Data Layer",
      tech: "React Context API (CartContext) & LocalStorage",
      desc: "Decoupled global cart state, persistent local storage for cart, orders & feedback"
    },
    {
      domain: "Development Environment",
      tech: "Antigravity AI Development Environment",
      desc: "Google DeepMind pair-programming agentic development environment"
    },
    {
      domain: "Version Control",
      tech: "Git & GitHub",
      desc: "Distributed version control system managing branches, commits and continuous delivery"
    },
    {
      domain: "Production Deployment",
      tech: "Vercel Platform",
      desc: "High-performance edge CDN static hosting with automated deployment pipeline"
    }
  ];

  const headerRow = new TableRow({
    children: [
      new TableCell({
        width: { size: 28, type: WidthType.PERCENTAGE },
        shading: { fill: "EEF2FF", type: ShadingType.CLEAR },
        children: [
          new Paragraph({
            children: [new TextRun({ text: "Technology Domain", bold: true, size: 18, color: COLOR_PRIMARY_DARK, font: "Calibri" })],
            spacing: { before: 40, after: 40 },
          }),
        ],
      }),
      new TableCell({
        width: { size: 36, type: WidthType.PERCENTAGE },
        shading: { fill: "EEF2FF", type: ShadingType.CLEAR },
        children: [
          new Paragraph({
            children: [new TextRun({ text: "Technology / Framework & Version", bold: true, size: 18, color: COLOR_PRIMARY_DARK, font: "Calibri" })],
            spacing: { before: 40, after: 40 },
          }),
        ],
      }),
      new TableCell({
        width: { size: 36, type: WidthType.PERCENTAGE },
        shading: { fill: "EEF2FF", type: ShadingType.CLEAR },
        children: [
          new Paragraph({
            children: [new TextRun({ text: "Role in Architecture", bold: true, size: 18, color: COLOR_PRIMARY_DARK, font: "Calibri" })],
            spacing: { before: 40, after: 40 },
          }),
        ],
      }),
    ],
  });

  const rows = [headerRow];

  tableData.forEach((item, index) => {
    const bg = index % 2 === 0 ? "FFFFFF" : COLOR_LIGHT_BG;
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            width: { size: 28, type: WidthType.PERCENTAGE },
            shading: { fill: bg, type: ShadingType.CLEAR },
            children: [
              new Paragraph({
                children: [new TextRun({ text: item.domain, bold: true, size: 17, color: COLOR_DARK, font: "Calibri" })],
                spacing: { before: 25, after: 25 },
              }),
            ],
          }),
          new TableCell({
            width: { size: 36, type: WidthType.PERCENTAGE },
            shading: { fill: bg, type: ShadingType.CLEAR },
            children: [
              new Paragraph({
                children: [new TextRun({ text: item.tech, bold: false, size: 17, color: COLOR_PRIMARY, font: "Calibri" })],
                spacing: { before: 25, after: 25 },
              }),
            ],
          }),
          new TableCell({
            width: { size: 36, type: WidthType.PERCENTAGE },
            shading: { fill: bg, type: ShadingType.CLEAR },
            children: [
              new Paragraph({
                children: [new TextRun({ text: item.desc, size: 16, color: COLOR_MUTED, font: "Calibri" })],
                spacing: { before: 25, after: 25 },
              }),
            ],
          }),
        ],
      })
    );
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows,
  });
}

function createFeaturesGrid() {
  const features = [
    { title: "Responsive Design", desc: "Mobile-first grid & flex layouts tested on mobile, tablet, and desktop viewports." },
    { title: "Product Catalogue", desc: "25 products with INR (₹) pricing, stock statuses, discount tags, and deep specifications." },
    { title: "Product Search & Filter", desc: "Instant real-time search across titles/descriptions, category filtering, and multi-criteria sorting." },
    { title: "Shopping Cart & Checkout", desc: "Interactive slide-over cart, quantity controls, coupon engine, and simulated payment confirmation." },
    { title: "Product Details View", desc: "Modal window with 4K image preview, technical spec sheet, feature bullets, and stock badges." },
    { title: "Business Model Canvas", desc: "Complete 9-component Osterwalder canvas tailored to TechNest's D2C electronics strategy." },
    { title: "Marketing Strategy", desc: "9 digital marketing techniques, promotional campaign dashboards, and full 7-stage conversion funnel." },
    { title: "Order Fulfillment Process", desc: "10-stage visual order-to-delivery journey, exception paths, and 6 academic operational KPIs." },
  ];

  const rows = [];
  for (let i = 0; i < features.length; i += 2) {
    const f1 = features[i];
    const f2 = features[i + 1];

    rows.push(
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                spacing: { before: 20, after: 20 },
                children: [
                  new TextRun({ text: `• ${f1.title}: `, bold: true, size: 17, color: COLOR_PRIMARY, font: "Calibri" }),
                  new TextRun({ text: f1.desc, size: 16, color: COLOR_DARK, font: "Calibri" }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                spacing: { before: 20, after: 20 },
                children: [
                  new TextRun({ text: `• ${f2.title}: `, bold: true, size: 17, color: COLOR_PRIMARY, font: "Calibri" }),
                  new TextRun({ text: f2.desc, size: 16, color: COLOR_DARK, font: "Calibri" }),
                ],
              }),
            ],
          }),
        ],
      })
    );
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: rows,
  });
}

function createDeploymentSection() {
  return new Paragraph({
    spacing: { before: 40, after: 60 },
    alignment: AlignmentType.JUSTIFY,
    children: [
      new TextRun({
        text: "Deployment Platform: ",
        bold: true,
        size: 18,
        color: COLOR_PRIMARY_DARK,
        font: "Calibri",
      }),
      new TextRun({
        text: "The application is built using Vite's optimized production bundler (`vite build`), outputting minified static HTML, CSS, and JavaScript bundles. The compiled artifacts are configured for zero-configuration, continuous edge deployment on the ",
        size: 18,
        color: COLOR_DARK,
        font: "Calibri",
      }),
      new TextRun({
        text: "Vercel Platform",
        bold: true,
        size: 18,
        color: COLOR_ACCENT,
        font: "Calibri",
      }),
      new TextRun({
        text: " with automatic Single Page Application (SPA) routing rules (`vercel.json`), providing global CDN acceleration, automatic SSL certification, and 99.99% uptime for academic evaluation.",
        size: 18,
        color: COLOR_DARK,
        font: "Calibri",
      }),
    ],
  });
}

function createFooterNote() {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 0 },
    children: [
      new TextRun({
        text: "TechNest Store © 2026 | Prepared for MBA E-Commerce Academic Evaluation | Certified 1-Page Summary",
        italics: true,
        size: 15, // 7.5pt
        color: COLOR_MUTED,
        font: "Calibri",
      }),
    ],
  });
}

async function generateDocx() {
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 576,    // 0.4 inch (in twips: 1440 * 0.4 = 576)
              bottom: 576, // 0.4 inch
              left: 720,   // 0.5 inch (720 twips)
              right: 720,  // 0.5 inch
            },
          },
        },
        children: [
          createHeaderBanner(),
          createSectionHeading("1. Project Overview"),
          createParagraph(
            "TechNest is an academic e-commerce web application developed to demonstrate a contemporary direct-to-consumer (D2C) online retail business model, interactive product catalogue, comprehensive digital marketing strategy, and end-to-end order fulfillment pipeline. Built specifically for an MBA E-Commerce curriculum, the application combines commercial-grade user experience with academic rigor, illustrating essential multi-channel sales mechanics, consumer sentiment analysis, and lean startup frameworks."
          ),
          createSectionHeading("2. Framework & Technology Stack (Actual Project Technologies)"),
          createTechTable(),
          createSectionHeading("3. Key Features"),
          createFeaturesGrid(),
          createSectionHeading("4. Deployment Architecture"),
          createDeploymentSection(),
          createFooterNote(),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  
  // Write to project root
  const rootPath = path.resolve(__dirname, '../TechNest_Framework_and_Technology.docx');
  fs.writeFileSync(rootPath, buffer);
  console.log(`Saved Word document to: ${rootPath}`);

  // Write to public folder for web download
  const publicPath = path.resolve(__dirname, '../public/TechNest_Framework_and_Technology.docx');
  fs.writeFileSync(publicPath, buffer);
  console.log(`Saved Word document to: ${publicPath}`);
}

generateDocx().catch(err => {
  console.error("Error generating docx:", err);
  process.exit(1);
});
