/**
 * Schofield Construction — RAMS Generator
 * Single Node.js app that serves the HTML page and handles the AI calls.
 *
 * Set the ANTHROPIC_API_KEY environment variable in Hostinger's
 * Node.js app settings before starting.
 */

require('dotenv').config();
const express = require('express');
const path = require('path');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
app.use(cors());
app.use(express.json({ limit: '20mb' }));

// Serve the HTML page and any other static files from the "public" folder
app.use(express.static(path.join(__dirname, 'public')));

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// Simple rate limit so a leaked endpoint can't burn through credits
const callLog = new Map();
function rateLimit(req, res, next) {
  const ip = req.ip;
  const now = Date.now();
  const bucket = callLog.get(ip) || [];
  const recent = bucket.filter(t => now - t < 60_000);
  if (recent.length >= 10) {
    return res.status(429).json({ error: 'Too many requests, please wait a minute.' });
  }
  recent.push(now);
  callLog.set(ip, recent);
  next();
}

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
- Electrical hazards (live services, isolation, LOTO)
- Plant & equipment (PUWER), lifting operations (LOLER)
- Hazardous substances (COSHH) — dust, silica, solvents, sealants
- Asbestos (CAR 2012) — assume pre-2000 buildings need R&D survey
- Confined spaces, excavations
- Fire and emergency
- Noise and HAVS
- Welfare, COVID/biological where relevant
- Public, occupants and adjacent contractors interface
- Traffic management on site
- Temporary works

# JSON schema (return EXACTLY this structure)
{
  "title": "Risk Assessment & Method Statement — <Project>",
  "project": {
    "name": string,
    "address": string,
    "contractor": string,
    "principalContractor": string,
    "docRef": string (always set to "N/A" — the contractor will fill in their own reference at issue),
    "dateIssued": string (e.g. "10 May 2026"),
    "reviewDate": string (90 days after issue, or sooner if scope changes)
  },
  "scope": string (reworded clean version of contractor's scope, 1-3 paragraphs),
  "hazards": [
    {
      "activity": string (specific work activity),
      "hazard": string (what causes harm),
      "personsAtRisk": [string, ...],
      "likelihood": integer 1-5,
      "severity": integer 1-5,
      "initialRisk": integer (likelihood * severity),
      "controls": [string, ...],
      "residualRisk": integer
    }
  ],
  "ppe": [string, ...],
  "method": [
    { "step": string, "detail": string }
  ],
  "plantEquipment": [string, ...],
  "coshh": [
    { "substance": string, "hazard": string, "controls": string }
  ],
  "emergency": [string, ...],
  "training": [string, ...],
  "assumptions": [string, ...]
}

# Critical rules
1. Be specific to the trade and scope. No generic boilerplate.
2. Identify at least 5 distinct hazards for any non-trivial scope, more if warranted.
3. Where the supplied documentation is silent on a critical point, add a clear item to "assumptions" rather than inventing facts.
4. Use UK terminology.
5. Cite EN/BS standards for PPE where applicable.
6. The method statement steps must be in logical sequence.
7. ALL output keys are required. If a section genuinely doesn't apply, return [] not omitted.
8. Output JSON only. Nothing else.`;

app.post('/api/generate-rams', rateLimit, async (req, res) => {
  try {
    const { project, docText } = req.body || {};
    if (!project || !project.projectName || !project.scope) {
      return res.status(400).json({ error: 'projectName and scope are required' });
    }

    const userMessage = `# Project information supplied by contractor

Project name: ${project.projectName}
Site address: ${project.siteAddress || '[not supplied]'}
Contractor: ${project.contractor || '[not supplied]'}
Principal Contractor: ${project.principalContractor || '[same as contractor]'}

# Scope of works
${project.scope}

# Known hazards / constraints flagged by contractor
${project.knownHazards || '[none flagged — you must still identify hazards from the scope]'}

# Supporting site documentation (extracted text)
${docText && docText.length > 50
  ? docText
  : '[No supporting documents uploaded. Generate the RAMS from the scope and call out gaps in the "assumptions" array.]'}

---
Produce the JSON RAMS now. Output JSON only.`;

    const response = await client.messages.create({
      model: 'claude-opus-4-5',
      max_tokens: 8000,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userMessage }]
    });

    let raw = response.content
      .filter(b => b.type === 'text')
      .map(b => b.text)
      .join('\n')
      .trim();

    raw = raw.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '').trim();

    let rams;
    try {
      rams = JSON.parse(raw);
    } catch (err) {
      console.error('Failed to parse model output as JSON:', err);
      return res.status(502).json({ error: 'AI returned invalid JSON' });
    }

    rams.project = rams.project || {};
    rams.project.docRef = 'N/A';
    rams.project.dateIssued = rams.project.dateIssued || new Date().toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' });

    res.json(rams);
  } catch (err) {
    console.error('RAMS generation error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Health check
app.get('/api/health', (req, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Schofield RAMS server listening on :${PORT}`);
});
