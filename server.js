/**
 * Schofield Construction — RAMS Generator
 * v1.3 — more robust Anthropic handling, safer JSON parsing, better Render stability
 *
 * Required environment variables:
 *   ANTHROPIC_API_KEY — Claude API key for AI generation
 *   RESEND_API_KEY    — Resend API key for sending emails
 *
 * Optional environment variables:
 *   ANTHROPIC_MODEL   — defaults to claude-opus-4-5
 */

require('dotenv').config();

const express = require('express');
const path = require('path');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const { Resend } = require('resend');

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  BorderStyle,
  WidthType,
  ShadingType,
  LevelFormat,
  PageNumber,
  Header,
  Footer
} = require('docx');

const app = express();

app.use(cors());
app.use(express.json({ limit: '12mb' }));
app.use(express.static(path.join(__dirname, 'public')));

if (!process.env.ANTHROPIC_API_KEY) {
  console.warn('WARNING: ANTHROPIC_API_KEY is missing.');
}

if (!process.env.RESEND_API_KEY) {
  console.warn('WARNING: RESEND_API_KEY is missing.');
}

const claude = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY
});

const resend = new Resend(process.env.RESEND_API_KEY);

const ANTHROPIC_MODEL = process.env.ANTHROPIC_MODEL || 'claude-opus-4-5';

// Internal CC address — always copied on RAMS emails
const INTERNAL_CC = 'info@schofieldconstruction.site';
const FROM_ADDRESS = 'Schofield Construction <info@schofieldconstruction.site>';

// Keep this lower than the frontend 80,000 char limit.
// This protects Render + Anthropic from very large uploaded docs.
const MAX_DOC_TEXT_CHARS = 25000;
const MAX_SCOPE_CHARS = 6000;
const MAX_HAZARDS_CHARS = 4000;

// ===========================================================
// SIMPLE RATE LIMIT
// ===========================================================
const callLog = new Map();

function rateLimit(req, res, next) {
  const ip = req.ip || req.headers['x-forwarded-for'] || 'unknown';
  const now = Date.now();
  const bucket = callLog.get(ip) || [];
  const recent = bucket.filter(t => now - t < 60_000);

  if (recent.length >= 10) {
    return res.status(429).json({
      error: 'Too many requests, please wait a minute.'
    });
  }

  recent.push(now);
  callLog.set(ip, recent);
  next();
}

// ===========================================================
// IDEMPOTENCY CACHE
// Prevents duplicate emails when the frontend retries.
// ===========================================================
const idempotencyCache = new Map();
const inFlight = new Map();
const IDEMPOTENCY_TTL = 5 * 60 * 1000; // 5 minutes

function getIdempotent(key) {
  if (!key) return null;

  const entry = idempotencyCache.get(key);
  if (!entry) return null;

  if (Date.now() - entry.savedAt > IDEMPOTENCY_TTL) {
    idempotencyCache.delete(key);
    return null;
  }

  return entry.response;
}

function saveIdempotent(key, response) {
  if (!key) return;

  idempotencyCache.set(key, {
    savedAt: Date.now(),
    response
  });

  if (idempotencyCache.size > 1000) {
    const cutoff = Date.now() - IDEMPOTENCY_TTL;

    for (const [k, v] of idempotencyCache) {
      if (v.savedAt < cutoff) {
        idempotencyCache.delete(k);
      }
    }
  }
}

function truncateText(value, maxChars) {
  const text = String(value || '');
  if (text.length <= maxChars) return text;

  return text.slice(0, maxChars) +
    `\n\n[TRUNCATED: original extracted text was ${text.length} characters. Only the first ${maxChars} characters were sent for stability.]`;
}

function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ===========================================================
// SYSTEM PROMPT
// ===========================================================
const SYSTEM_PROMPT = `You are a UK Construction Health & Safety expert producing a Risk Assessment & Method Statement (RAMS) document compliant with the Construction (Design and Management) Regulations 2015 (CDM 2015) and HSE guidance HSG150.

Your output MUST be valid JSON conforming exactly to the schema described below. Do not include any explanatory prose outside the JSON. Do not wrap the JSON in markdown code fences.

# Risk scoring methodology
Use the standard HSE 5x5 matrix:
- Likelihood: 1=Very rare, 2=Unlikely, 3=Possible, 4=Likely, 5=Almost certain
- Severity:   1=Minor first aid, 2=Lost time injury, 3=>7-day RIDDOR, 4=Major injury, 5=Fatality
- Risk = Likelihood x Severity (1-25)
- Apply the hierarchy of control: Eliminate -> Substitute -> Engineering controls -> Administrative controls -> PPE
- "Initial" risk is BEFORE controls. "Residual" is AFTER controls. Residual MUST be lower than initial.
- Residual risk should be reduced to "Low" (<=7) wherever practicable.

# Coverage requirements
For the scope provided, you must identify and assess every reasonably foreseeable hazard. Standard construction hazards to consider:
- Manual handling, working at height, falling objects
- Slips, trips, falls on the level
- Electrical hazards, live services, isolation, LOTO
- Plant & equipment, PUWER, lifting operations, LOLER
- Hazardous substances, COSHH, dust, silica, solvents, sealants
- Asbestos, CAR 2012, assume pre-2000 buildings need R&D survey
- Confined spaces and excavations where relevant
- Fire and emergency
- Noise and HAVS
- Welfare
- Public, occupants and adjacent contractors interface
- Traffic management on site
- Temporary works

# JSON schema
Return EXACTLY this structure:

{
  "title": "Risk Assessment & Method Statement — <Project>",
  "project": {
    "name": string,
    "address": string,
    "contractor": string,
    "principalContractor": string,
    "docRef": "N/A",
    "dateIssued": string,
    "reviewDate": ""
  },
  "scope": string,
  "hazards": [
    {
      "activity": string,
      "hazard": string,
      "personsAtRisk": [string],
      "likelihood": integer,
      "severity": integer,
      "initialRisk": integer,
      "controls": [string],
      "residualRisk": integer
    }
  ],
  "ppe": [string],
  "method": [
    {
      "step": string,
      "detail": string
    }
  ],
  "plantEquipment": [string],
  "coshh": [
    {
      "substance": string,
      "hazard": string,
      "controls": string
    }
  ],
  "emergency": [string],
  "training": [string],
  "assumptions": [string]
}

# Critical rules
1. Be specific to the trade and scope. No generic boilerplate.
2. Identify at least 5 distinct hazards for any non-trivial scope.
3. Where the supplied documentation is silent on a critical point, add a clear item to "assumptions".
4. Use UK terminology.
5. Cite EN/BS standards for PPE where applicable.
6. The method statement steps must be in logical sequence.
7. ALL output keys are required. If a section genuinely does not apply, return [] not omitted.
8. Output JSON only. Nothing else.
9. Keep the JSON compact enough to complete reliably. Do not write unnecessary long prose.`;

// ===========================================================
// JSON HELPERS
// ===========================================================
function extractJson(text) {
  let s = String(text || '').trim();

  s = s
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/```\s*$/i, '')
    .trim();

  const firstBrace = s.indexOf('{');
  const lastBrace = s.lastIndexOf('}');

  if (firstBrace !== -1 && lastBrace > firstBrace) {
    s = s.substring(firstBrace, lastBrace + 1);
  }

  return s;
}

function safeArray(value) {
  if (Array.isArray(value)) return value;
  if (value === undefined || value === null || value === '') return [];
  return [String(value)];
}

function safeString(value, fallback = '') {
  if (value === undefined || value === null) return fallback;
  return String(value);
}

function clampInt(value, min, max, fallback) {
  const n = Number.parseInt(value, 10);
  if (Number.isNaN(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

function normaliseRams(rams, projectInput) {
  const today = new Date().toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  });

  const output = rams && typeof rams === 'object' ? rams : {};

  output.project = output.project && typeof output.project === 'object'
    ? output.project
    : {};

  output.title = safeString(
    output.title,
    `Risk Assessment & Method Statement — ${projectInput.projectName || 'Project'}`
  );

  output.project.name = safeString(output.project.name, projectInput.projectName || 'Project');
  output.project.address = safeString(output.project.address, projectInput.siteAddress || 'Not supplied');
  output.project.contractor = safeString(output.project.contractor, projectInput.contractor || 'Not supplied');
  output.project.principalContractor = safeString(
    output.project.principalContractor,
    projectInput.principalContractor || projectInput.contractor || 'Not supplied'
  );

  output.project.docRef = 'N/A';
  output.project.dateIssued = today;
  output.project.reviewDate = '';

  output.scope = safeString(output.scope, projectInput.scope || '');

  output.hazards = safeArray(output.hazards).map((h, index) => {
    const item = h && typeof h === 'object' ? h : {};

    const likelihood = clampInt(item.likelihood, 1, 5, 3);
    const severity = clampInt(item.severity, 1, 5, 3);
    const initialRisk = clampInt(item.initialRisk, 1, 25, likelihood * severity);

    let residualRisk = clampInt(item.residualRisk, 1, 25, Math.max(1, Math.min(7, initialRisk - 2)));
    if (residualRisk >= initialRisk) {
      residualRisk = Math.max(1, initialRisk - 1);
    }

    return {
      activity: safeString(item.activity, `Activity ${index + 1}`),
      hazard: safeString(item.hazard, 'Hazard to be reviewed'),
      personsAtRisk: safeArray(item.personsAtRisk).map(String),
      likelihood,
      severity,
      initialRisk,
      controls: safeArray(item.controls).map(String),
      residualRisk
    };
  });

  output.ppe = safeArray(output.ppe).map(String);

  output.method = safeArray(output.method).map((m, index) => {
    const item = m && typeof m === 'object' ? m : {};
    return {
      step: safeString(item.step, `Step ${index + 1}`)
        .replace(/^\s*(?:step\s*)?\d+[\.\):\-]\s*/i, '')
        .trim(),
      detail: safeString(item.detail, String(m || ''))
    };
  });

  output.plantEquipment = safeArray(output.plantEquipment).map(String);

  output.coshh = safeArray(output.coshh).map(c => {
    const item = c && typeof c === 'object' ? c : {};
    return {
      substance: safeString(item.substance, 'To be confirmed'),
      hazard: safeString(item.hazard, 'To be reviewed'),
      controls: safeString(item.controls, 'Use in accordance with manufacturer instructions and COSHH assessment.')
    };
  });

  output.emergency = safeArray(output.emergency).map(String);
  output.training = safeArray(output.training).map(String);
  output.assumptions = safeArray(output.assumptions).map(String);

  if (!output.hazards.length) {
    output.hazards.push({
      activity: 'General construction activity',
      hazard: 'Hazards not fully identified due to limited information supplied',
      personsAtRisk: ['Operatives', 'Other contractors', 'Members of the public where applicable'],
      likelihood: 3,
      severity: 3,
      initialRisk: 9,
      controls: [
        'Competent person to review the scope before works commence.',
        'Site-specific briefing to be completed before starting.',
        'Stop works if undocumented hazards are identified.'
      ],
      residualRisk: 4
    });

    output.assumptions.push('Limited information was supplied. The RAMS must be reviewed by a competent person before issue.');
  }

  if (!output.method.length) {
    output.method.push({
      step: 'Pre-start review',
      detail: 'Review the site conditions, scope, access, emergency arrangements and known hazards before works commence.'
    });
  }

  return output;
}

function parseRamsJson(rawText, projectInput) {
  const cleaned = extractJson(rawText);

  try {
    const parsed = JSON.parse(cleaned);
    return normaliseRams(parsed, projectInput);
  } catch (err) {
    console.error('JSON parse failed:', err.message);
    console.error('Raw AI output preview:', String(rawText || '').slice(0, 1200));
    throw new Error('AI returned invalid JSON.');
  }
}

// ===========================================================
// ANTHROPIC CALL
// ===========================================================
async function callClaudeForRams(project, docText) {
  const safeProject = {
    projectName: truncateText(project.projectName, 300),
    siteAddress: truncateText(project.siteAddress, 500),
    contractor: truncateText(project.contractor, 300),
    principalContractor: truncateText(project.principalContractor, 300),
    scope: truncateText(project.scope, MAX_SCOPE_CHARS),
    knownHazards: truncateText(project.knownHazards, MAX_HAZARDS_CHARS)
  };

  const safeDocText = truncateText(docText || '', MAX_DOC_TEXT_CHARS);

  const userMessage = `# Project information supplied by contractor

Project name: ${safeProject.projectName}
Site address: ${safeProject.siteAddress || '[not supplied]'}
Contractor: ${safeProject.contractor || '[not supplied]'}
Principal Contractor: ${safeProject.principalContractor || '[same as contractor]'}

# Scope of works
${safeProject.scope}

# Known hazards / constraints flagged by contractor
${safeProject.knownHazards || '[none flagged — you must still identify hazards from the scope]'}

# Supporting site documentation, extracted text
${safeDocText && safeDocText.length > 50
    ? safeDocText
    : '[No supporting documents uploaded. Generate the RAMS from the scope and call out gaps in the assumptions array.]'}

---
Produce the JSON RAMS now. Output JSON only.`;

  const attempts = [
    {
      name: 'primary',
      maxTokens: 5000,
      extraInstruction: ''
    },
    {
      name: 'json-retry',
      maxTokens: 4500,
      extraInstruction: 'CRITICAL: Return ONLY a valid JSON object. No markdown. No commentary. Start with { and end with }. Keep wording concise.'
    },
    {
      name: 'compact-retry',
      maxTokens: 3500,
      extraInstruction: 'CRITICAL: Return compact valid JSON only. Reduce wording length. Include at least 5 hazards but keep each control concise.'
    }
  ];

  let lastError = null;

  for (let i = 0; i < attempts.length; i++) {
    const attempt = attempts[i];

    try {
      console.log(`Calling Anthropic attempt ${i + 1}/${attempts.length}: ${attempt.name}`);

      const response = await claude.messages.create({
        model: ANTHROPIC_MODEL,
        max_tokens: attempt.maxTokens,
        temperature: 0.2,
        system: SYSTEM_PROMPT,
        messages: [
          {
            role: 'user',
            content: userMessage + (attempt.extraInstruction ? `\n\n${attempt.extraInstruction}` : '')
          }
        ]
      });

      const raw = response.content
        .filter(block => block.type === 'text')
        .map(block => block.text)
        .join('\n')
        .trim();

      if (!raw) {
        throw new Error('Anthropic returned an empty response.');
      }

      const rams = parseRamsJson(raw, project);
      console.log(`Anthropic attempt ${attempt.name} succeeded.`);
      return rams;

    } catch (err) {
      lastError = err;

      const message = err && err.message ? err.message : String(err);
      console.error(`Anthropic attempt ${attempt.name} failed:`, message);

      if (i < attempts.length - 1) {
        await wait(1500 * (i + 1));
      }
    }
  }

  const finalMessage = lastError && lastError.message
    ? lastError.message
    : 'Unknown Anthropic failure';

  throw new Error(`AI generation failed after retries: ${finalMessage}`);
}

// ============================================================
// WORD DOCUMENT BUILDER
// ============================================================
function buildRamsDocx(r) {
  const border = {
    style: BorderStyle.SINGLE,
    size: 6,
    color: '1A1F1B'
  };

  const cellBorders = {
    top: border,
    bottom: border,
    left: border,
    right: border
  };

  const headerCell = (text, width) => new TableCell({
    borders: cellBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: '1A1F1B', type: ShadingType.CLEAR, color: 'auto' },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text,
            bold: true,
            color: 'FFD400',
            size: 18
          })
        ]
      })
    ]
  });

  const cell = (text, width, opts = {}) => new TableCell({
    borders: cellBorders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill
      ? { fill: opts.fill, type: ShadingType.CLEAR, color: 'auto' }
      : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: (Array.isArray(text) ? text : [text]).map(t =>
      new Paragraph({
        children: [
          new TextRun({
            text: String(t || ''),
            size: 20,
            bold: opts.bold || false
          })
        ]
      })
    )
  });

  const riskFill = score => {
    const n = Number(score) || 0;
    return n >= 15 ? 'FDE0DB' : n >= 8 ? 'FFF3CC' : 'DFF0D6';
  };

  const children = [];

  children.push(new Paragraph({
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [
      new TextRun({
        text: r.title || 'RISK ASSESSMENT & METHOD STATEMENT',
        bold: true,
        size: 40
      })
    ]
  }));

  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
    children: [
      new TextRun({
        text: 'CDM 2015 Compliant',
        italics: true,
        size: 22,
        color: '5A5F56'
      })
    ]
  }));

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [
      new TextRun({
        text: '1. Project Information',
        bold: true
      })
    ]
  }));

  const projRows = [
    ['Project', r.project.name],
    ['Site address', r.project.address],
    ['Contractor', r.project.contractor],
    ['Principal Contractor', r.project.principalContractor || '—'],
    ['Document ref', r.project.docRef || 'N/A'],
    ['Date issued', r.project.dateIssued],
    ['Review date', r.project.reviewDate || 'To be set on issue']
  ];

  children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2800, 6560],
    rows: projRows.map(([k, v]) => new TableRow({
      children: [
        cell(k, 2800, { fill: 'EBE6DA', bold: true }),
        cell(v || '—', 6560)
      ]
    }))
  }));

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300 },
    children: [
      new TextRun({
        text: '2. Scope of Works',
        bold: true
      })
    ]
  }));

  children.push(new Paragraph({
    children: [
      new TextRun({
        text: r.scope || '',
        size: 22
      })
    ]
  }));

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300 },
    children: [
      new TextRun({
        text: '3. Hazard Identification & Risk Assessment (5×5)',
        bold: true
      })
    ]
  }));

  children.push(new Paragraph({
    spacing: { after: 150 },
    children: [
      new TextRun({
        text: 'L = Likelihood (1–5), S = Severity (1–5), Risk = L × S. Green ≤7, Amber 8–14, Red ≥15.',
        size: 18,
        italics: true,
        color: '5A5F56'
      })
    ]
  }));

  children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [400, 1700, 1300, 400, 400, 600, 3360, 1200],
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          headerCell('#', 400),
          headerCell('Activity / hazard', 1700),
          headerCell('Persons at risk', 1300),
          headerCell('L', 400),
          headerCell('S', 400),
          headerCell('Initial', 600),
          headerCell('Controls (hierarchy)', 3360),
          headerCell('Residual', 1200)
        ]
      }),
      ...(r.hazards || []).map((h, i) => new TableRow({
        children: [
          cell(String(i + 1), 400),
          cell([h.activity, h.hazard], 1700),
          cell((h.personsAtRisk || []).join(', '), 1300),
          cell(String(h.likelihood), 400),
          cell(String(h.severity), 400),
          cell(String(h.initialRisk), 600, {
            fill: riskFill(h.initialRisk),
            bold: true
          }),
          cell((h.controls || []).map(c => `• ${c}`), 3360),
          cell(String(h.residualRisk), 1200, {
            fill: riskFill(h.residualRisk),
            bold: true
          })
        ]
      }))
    ]
  }));

  const bulletSection = (title, items, num) => {
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 300 },
      children: [
        new TextRun({
          text: `${num}. ${title}`,
          bold: true
        })
      ]
    }));

    if (items && items.length) {
      items.forEach(t => {
        children.push(new Paragraph({
          numbering: {
            reference: 'bullets',
            level: 0
          },
          children: [
            new TextRun({
              text: String(t),
              size: 22
            })
          ]
        }));
      });
    } else {
      children.push(new Paragraph({
        children: [
          new TextRun({
            text: 'Not applicable.',
            italics: true,
            size: 22
          })
        ]
      }));
    }
  };

  bulletSection('PPE Requirements', r.ppe, 4);

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300 },
    children: [
      new TextRun({
        text: '5. Method Statement',
        bold: true
      })
    ]
  }));

  (r.method || []).forEach(s => {
    const cleanStep = String(s.step || '')
      .replace(/^\s*(?:step\s*)?\d+[\.\):\-]\s*/i, '')
      .trim();

    children.push(new Paragraph({
      numbering: {
        reference: 'numbers',
        level: 0
      },
      children: [
        new TextRun({
          text: `${cleanStep || 'Step'}: `,
          bold: true,
          size: 22
        }),
        new TextRun({
          text: String(s.detail || ''),
          size: 22
        })
      ]
    }));
  });

  bulletSection('Plant, Equipment & Materials', r.plantEquipment, 6);

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300 },
    children: [
      new TextRun({
        text: '7. COSHH / Hazardous Substances',
        bold: true
      })
    ]
  }));

  if (r.coshh && r.coshh.length) {
    children.push(new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 2800, 4160],
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            headerCell('Substance', 2400),
            headerCell('Hazard', 2800),
            headerCell('Controls', 4160)
          ]
        }),
        ...r.coshh.map(c => new TableRow({
          children: [
            cell(c.substance, 2400),
            cell(c.hazard, 2800),
            cell(c.controls, 4160)
          ]
        }))
      ]
    }));
  } else {
    children.push(new Paragraph({
      children: [
        new TextRun({
          text: 'No hazardous substances identified for this scope.',
          italics: true,
          size: 22
        })
      ]
    }));
  }

  bulletSection('Emergency Procedures', r.emergency, 8);
  bulletSection('Training & Competence', r.training, 9);

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300 },
    children: [
      new TextRun({
        text: '10. Sign-off & Briefing',
        bold: true
      })
    ]
  }));

  const signRows = [
    ['Prepared by', ''],
    ['Position', ''],
    ['Signature / Date', ''],
    ['Reviewed by (PC)', ''],
    ['Briefed to operatives', 'see attached briefing register']
  ];

  children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2800, 6560],
    rows: signRows.map(([k, v]) => new TableRow({
      children: [
        cell(k, 2800, { fill: 'EBE6DA', bold: true }),
        cell(v, 6560)
      ]
    }))
  }));

  if (r.assumptions && r.assumptions.length) {
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 300 },
      children: [
        new TextRun({
          text: 'Assumptions & Gaps Flagged for Review',
          bold: true,
          color: 'C8341C'
        })
      ]
    }));

    r.assumptions.forEach(a => {
      children.push(new Paragraph({
        numbering: {
          reference: 'bullets',
          level: 0
        },
        children: [
          new TextRun({
            text: a,
            size: 22
          })
        ]
      }));
    });
  }

  const doc = new Document({
    creator: 'Schofield Construction RAMS Generator',
    title: r.title,
    styles: {
      default: {
        document: {
          run: {
            font: 'Calibri',
            size: 22
          }
        }
      },
      paragraphStyles: [
        {
          id: 'Title',
          name: 'Title',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: 40,
            bold: true,
            font: 'Calibri'
          },
          paragraph: {
            spacing: {
              before: 0,
              after: 200
            },
            outlineLevel: 0
          }
        },
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: 28,
            bold: true,
            font: 'Calibri',
            color: '1A1F1B'
          },
          paragraph: {
            spacing: {
              before: 240,
              after: 120
            },
            outlineLevel: 0
          }
        }
      ]
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: 540,
                    hanging: 270
                  }
                }
              }
            }
          ]
        },
        {
          reference: 'numbers',
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: '%1.',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: 540,
                    hanging: 270
                  }
                }
              }
            }
          ]
        }
      ]
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: 11906,
              height: 16838
            },
            margin: {
              top: 1134,
              right: 1134,
              bottom: 1134,
              left: 1134
            }
          }
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: 'RAMS drafted by Schofield Construction - www.schofieldconstruction.site/RAMS',
                    size: 16,
                    color: '5A5F56'
                  })
                ]
              })
            ]
          })
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: 'Page ',
                    size: 16,
                    color: '5A5F56'
                  }),
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    size: 16,
                    color: '5A5F56'
                  }),
                  new TextRun({
                    text: ' of ',
                    size: 16,
                    color: '5A5F56'
                  }),
                  new TextRun({
                    children: [PageNumber.TOTAL_PAGES],
                    size: 16,
                    color: '5A5F56'
                  })
                ]
              })
            ]
          })
        },
        children
      }
    ]
  });

  return Packer.toBuffer(doc);
}

// ============================================================
// EMAIL SENDER
// ============================================================
async function sendRamsEmail(contractorEmail, rams, docxBuffer) {
  const safeName = (rams.project.name || 'RAMS')
    .replace(/[^a-z0-9]+/gi, '_')
    .slice(0, 40);

  const stamp = new Date().toISOString().slice(0, 10);
  const filename = `${safeName}_RAMS_${stamp}.docx`;

  const hazardCount = (rams.hazards || []).length;
  const projectName = rams.project.name || 'your project';

  const html = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #0f1410;">
      <div style="background: #0f1410; color: #ffd400; padding: 20px; text-align: center;">
        <h1 style="margin: 0; font-size: 22px;">Schofield Construction</h1>
        <p style="margin: 4px 0 0; font-size: 13px; letter-spacing: 0.1em; text-transform: uppercase;">RAMS Generator</p>
      </div>
      <div style="padding: 24px; background: #f4f1ea;">
        <h2 style="font-size: 20px; margin-top: 0;">Your draft RAMS is ready</h2>
        <p>Hi,</p>
        <p>Attached is your draft Risk Assessment & Method Statement for <strong>${escapeHtml(projectName)}</strong>, generated via our online tool.</p>
        <p>The document covers <strong>${hazardCount} identified hazards</strong>, with full 5×5 risk scoring, hierarchy-of-control measures, sequential method statement, PPE requirements, COSHH register, and emergency procedures — aligned to CDM 2015 and HSG150.</p>
        <div style="background: #fff3cc; border-left: 4px solid #c8341c; padding: 12px 16px; margin: 20px 0;">
          <strong>Important:</strong> This is a draft for review. Please ensure a competent person reviews, amends as necessary, and signs off the RAMS before any works commence. Operatives must be briefed prior to starting on site.
        </div>
        <p>If you need to regenerate or amend the RAMS, you can return to <a href="https://schofieldconstruction.site/RAMS" style="color: #c8341c;">schofieldconstruction.site/RAMS</a> at any time.</p>
        <p style="margin-top: 30px;">Regards,<br><strong>Schofield Construction</strong></p>
      </div>
      <div style="background: #ebe6da; padding: 12px 24px; font-size: 11px; color: #5a5f56; text-align: center;">
        This RAMS was drafted using our AI-assisted tool. Final responsibility for safety on site rests with the Principal Contractor and the appointed competent person.
      </div>
    </div>
  `;

  const text = `Schofield Construction — Your draft RAMS is ready

Hi,

Attached is your draft Risk Assessment & Method Statement for ${projectName}, generated via our online tool.

The document covers ${hazardCount} identified hazards, with full 5x5 risk scoring, hierarchy-of-control measures, sequential method statement, PPE requirements, COSHH register, and emergency procedures — aligned to CDM 2015 and HSG150.

IMPORTANT: This is a draft for review. Please ensure a competent person reviews, amends as necessary, and signs off the RAMS before any works commence. Operatives must be briefed prior to starting on site.

If you need to regenerate or amend the RAMS, return to https://schofieldconstruction.site/RAMS

Regards,
Schofield Construction`;

  return resend.emails.send({
    from: FROM_ADDRESS,
    to: contractorEmail,
    cc: INTERNAL_CC,
    reply_to: INTERNAL_CC,
    subject: `Your RAMS — ${projectName}`,
    html,
    text,
    attachments: [
      {
        filename,
        content: docxBuffer
      }
    ]
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));
}

function isValidEmail(s) {
  if (!s || typeof s !== 'string') return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s.trim());
}

// ============================================================
// MAIN ENDPOINT
// ============================================================
app.post('/api/generate-rams', rateLimit, async (req, res) => {
  const startedAt = Date.now();
  const idempotencyKey = req.headers['x-idempotency-key'] || null;

  try {
    const { project, docText, contractorEmail } = req.body || {};

    if (!project || !project.projectName || !project.scope) {
      return res.status(400).json({
        error: 'Project name and scope are required.'
      });
    }

    if (!isValidEmail(contractorEmail)) {
      return res.status(400).json({
        error: 'A valid contractor email is required.'
      });
    }

    if (idempotencyKey) {
      const cached = getIdempotent(idempotencyKey);

      if (cached) {
        console.log('Idempotent hit:', idempotencyKey);
        return res.json(cached);
      }

      const pending = inFlight.get(idempotencyKey);

      if (pending) {
        console.log('Idempotent in-flight hit:', idempotencyKey);
        try {
          const result = await pending;
          return res.json(result);
        } catch (err) {
          return res.status(500).json({
            error: 'Request failed while already in progress. Please try again.'
          });
        }
      }
    }

    const workPromise = (async () => {
      console.log('RAMS generation started:', {
        projectName: project.projectName,
        hasDocText: Boolean(docText && String(docText).length > 50),
        docTextChars: String(docText || '').length,
        model: ANTHROPIC_MODEL
      });

      let rams;

      try {
        rams = await callClaudeForRams(project, docText);
      } catch (err) {
        console.error('AI generation failed:', err);
        throw new Error(
          'The AI service failed while generating the RAMS. This is usually a temporary Anthropic/Render connection issue or the input is too large. Try again, or shorten uploaded documents.'
        );
      }

      let docxBuffer;

      try {
        docxBuffer = await buildRamsDocx(rams);
      } catch (err) {
        console.error('Failed to build Word document:', err);
        throw new Error('Generated RAMS but failed to build the Word document.');
      }

      let emailStatus = 'sent';

      try {
        const result = await sendRamsEmail(contractorEmail, rams, docxBuffer);
        console.log('Email sent OK:', result?.data?.id || result);
      } catch (err) {
        console.error('Email send failed:', err);
        emailStatus = 'failed';
      }

      const responsePayload = {
        rams,
        emailStatus,
        contractorEmail
      };

      if (idempotencyKey) {
        saveIdempotent(idempotencyKey, responsePayload);
      }

      console.log('RAMS generation completed:', {
        projectName: rams.project?.name,
        emailStatus,
        ms: Date.now() - startedAt
      });

      return responsePayload;
    })();

    if (idempotencyKey) {
      inFlight.set(idempotencyKey, workPromise);
    }

    try {
      const payload = await workPromise;
      return res.json(payload);
    } finally {
      if (idempotencyKey) {
        inFlight.delete(idempotencyKey);
      }
    }

  } catch (err) {
    console.error('RAMS endpoint error:', err);

    if (idempotencyKey) {
      inFlight.delete(idempotencyKey);
    }

    return res.status(502).json({
      error: err.message || 'RAMS generation failed.'
    });
  }
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({
    ok: true,
    version: '1.3',
    model: ANTHROPIC_MODEL,
    maxDocTextChars: MAX_DOC_TEXT_CHARS
  });
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Schofield RAMS server v1.3 listening on :${PORT}`);
});
