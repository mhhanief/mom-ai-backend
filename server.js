require('dotenv').config();
const express = require('express');
const cors = require('cors');
const axios = require('axios');
const bodyParser = require('body-parser');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, PageBreak, AlignmentType } = require('docx');

const app = express();
const PORT = process.env.PORT || 3001;
const CLAUDE_API_KEY = process.env.CLAUDE_API_KEY;

// Middleware
app.use(cors());
app.use(bodyParser.json({ limit: '10mb' }));
app.use(bodyParser.urlencoded({ limit: '10mb', extended: true }));

// ============================================
// ROBUST JSON EXTRACTOR
// ============================================
function extractJSON(text) {
  try { return JSON.parse(text); } catch {}

  let cleaned = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
  try { return JSON.parse(cleaned); } catch {}

  const match = cleaned.match(/\{[\s\S]*\}/);
  if (match) {
    try { return JSON.parse(match[0]); } catch {}
  }

  let depth = 0, start = -1, end = -1;
  for (let i = 0; i < cleaned.length; i++) {
    if (cleaned[i] === '{') { if (depth === 0) start = i; depth++; }
    if (cleaned[i] === '}') { depth--; if (depth === 0) { end = i; break; } }
  }
  if (start !== -1 && end !== -1) {
    try { return JSON.parse(cleaned.slice(start, end + 1)); } catch {}
  }

  throw new Error('Could not extract JSON from AI response');
}

// ============================================
// MARKDOWN STRIPPER (for Word & PDF plain text)
// ============================================
function stripMarkdown(text) {
  return text
    .replace(/#{1,6}\s+/g, '')
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/\*(.*?)\*/g, '$1')
    .replace(/`(.*?)`/g, '$1')
    .replace(/^>\s+/gm, '')
    .replace(/^[-*+]\s+/gm, '• ')
    .trim();
}

// ============================================
// SYSTEM PROMPTS
// ============================================

const MOM_SYSTEM_PROMPT = `Anda adalah Professional Meeting Minutes (MoM) Agent dengan expertise tinggi dalam corporate documentation dan business communication.

GUIDELINES:
1. TONE: Professional, formal, clear, concise
2. GRAMMAR: Gunakan Bahasa Indonesia formal atau English sesuai input
3. ACTION ITEMS: Format jelas - "**[PIC Name]** - [Action] - [Deadline]"

MARKDOWN FORMATTING RULES (wajib):
- Gunakan **bold** untuk nama, istilah penting, PIC
- Gunakan bullet points (- item) untuk list agenda, kesepakatan
- Gunakan numbering (1. item) untuk action items berurutan
- Gunakan ### untuk sub-heading jika diperlukan

JANGAN: Tambahkan informasi yang tidak ada di input. Ubah keputusan yang sudah ditetapkan.

CRITICAL: Response HARUS berupa pure JSON saja. Tidak boleh ada teks apapun di luar JSON object.`;

const PRD_SYSTEM_PROMPT = `Anda adalah Expert Product Requirement Document (PRD) Agent dengan deep expertise dalam product management.

GUIDELINES:
1. TONE: Professional, strategic, clear, stakeholder-ready
2. LANGUAGE: Bahasa Indonesia formal
3. COMPLETENESS: Semua sections harus well-developed

MARKDOWN FORMATTING RULES (wajib digunakan di setiap field):
- Gunakan **bold** untuk istilah penting, nama fitur, KPI
- Gunakan bullet points (- item) untuk list fitur, personas
- Gunakan numbering (1. item) untuk langkah-langkah atau prioritas
- Gunakan ### untuk sub-section heading
- Gunakan > untuk highlight atau catatan penting

JANGAN: Invent features yang bukan di input. Change core requirements.

CRITICAL: Response HARUS berupa pure JSON saja. Tidak boleh ada teks apapun di luar JSON object. Field values HARUS mengandung Markdown formatting.`;

const CHAT_SYSTEM_PROMPT = `Anda adalah AI Assistant bernama Archi yang membantu tim RDC BNI dalam membuat dokumen PRD dan Meeting Minutes dengan lebih cepat.

KARAKTERISTIK:
- Helpful, friendly, dan professional
- Gunakan Bahasa Indonesia informal tapi tetap profesional
- Fokus pada topik: PRD creation, Meeting Minutes, document best practices

TONE: Friendly, conversational. Use "aku" dan "kamu" untuk lebih personal.`;

// Route health check
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// ============================================
// MEETING MINUTES ENDPOINTS
// ============================================

app.post('/api/enhance-mom', async (req, res) => {
  try {
    const { judul, tanggal, lokasi, partisipan, agenda, pembahasan, kesepakatan, actionItems } = req.body;

    if (!judul || !tanggal || !lokasi || !partisipan || !agenda || !pembahasan || !kesepakatan || !actionItems) {
      return res.status(400).json({
        error: 'Missing required fields',
        required: ['judul', 'tanggal', 'lokasi', 'partisipan', 'agenda', 'pembahasan', 'kesepakatan', 'actionItems']
      });
    }

    const userPrompt = `Berikut adalah raw input Meeting Minutes:

JUDUL MEETING: ${judul}
TANGGAL: ${tanggal}
LOKASI: ${lokasi}
PARTISIPAN: ${partisipan}
AGENDA: ${agenda}
PEMBAHASAN: ${pembahasan}
KESEPAKATAN: ${kesepakatan}
ACTION ITEMS: ${actionItems}

Enhance dan polish menjadi professional Meeting Minutes. Gunakan Markdown formatting dalam nilai field.

Output HANYA JSON berikut tanpa teks apapun di luar JSON:
{
  "judul": "string",
  "tanggal": "string (format: DD Bulan YYYY)",
  "lokasi": "string",
  "partisipan": "string (bullet points untuk list peserta)",
  "agenda": "string (numbered list)",
  "pembahasan": "string (### sub-heading per agenda, bullet points untuk detail)",
  "kesepakatan": "string (numbered list atau bullet points)",
  "actionItems": "string (numbered list, **PIC** - Action - Deadline)"
}`;

    const claudeResponse = await axios.post(
      'https://api.anthropic.com/v1/messages',
      {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 2048,
        system: MOM_SYSTEM_PROMPT,
        messages: [{ role: 'user', content: userPrompt }]
      },
      {
        headers: {
          'x-api-key': CLAUDE_API_KEY,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        }
      }
    );

    const aiResponse = claudeResponse.data.content[0].text;
    const enhancedMoM = extractJSON(aiResponse);

    res.json({ success: true, data: enhancedMoM, timestamp: new Date().toISOString() });

  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to enhance MoM', message: error.message, details: error.response?.data || null });
  }
});

// ============================================
// PRD ENDPOINTS
// ============================================

app.post('/api/generate-prd', async (req, res) => {
  try {
    const {
      productName, initiativeType, featureFamily, release,
      regulatoryPermission, procurement, poLead, poMember, pmo,
      overview, problemStatement, targetUsers, keyFeatures, successMetrics, notes
    } = req.body;

    if (!productName || !overview || !problemStatement || !targetUsers || !keyFeatures || !successMetrics) {
      return res.status(400).json({ error: 'Missing required fields for PRD generation' });
    }

    const userPrompt = `Berikut adalah PRD input:

PRODUCT NAME: ${productName}
INITIATIVE TYPE: ${initiativeType}
FEATURE FAMILY: ${featureFamily}
RELEASE: ${release}
REGULATORY PERMISSION: ${regulatoryPermission}
PROCUREMENT: ${procurement}
PO Lead: ${poLead} | PO Member: ${poMember} | PMO: ${pmo}

OVERVIEW: ${overview}
PROBLEM STATEMENT: ${problemStatement}
TARGET USERS: ${targetUsers}
KEY FEATURES: ${keyFeatures}
SUCCESS METRICS: ${successMetrics}
NOTES: ${notes || 'None'}

Enhance menjadi comprehensive PRD. WAJIB gunakan Markdown formatting (bold, bullet, numbering, ###) dalam setiap field.

Output HANYA JSON berikut tanpa teks apapun di luar JSON:
{
  "overview": "string (paragraf executive summary, **bold** untuk istilah kunci)",
  "problemStatement": "string (numbered list untuk pain points, **bold** untuk istilah penting)",
  "targetUsers": "string (### per persona, bullet points untuk karakteristik dan use case)",
  "keyFeatures": "string (### per phase/kategori, numbered list per fitur, **bold** nama fitur)",
  "successMetrics": "string (### per kategori metrik, bullet points untuk detail, **bold** nama KPI)"
}`;

    const claudeResponse = await axios.post(
      'https://api.anthropic.com/v1/messages',
      {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 3000,
        system: PRD_SYSTEM_PROMPT,
        messages: [{ role: 'user', content: userPrompt }]
      },
      {
        headers: {
          'x-api-key': CLAUDE_API_KEY,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        }
      }
    );

    const aiResponse = claudeResponse.data.content[0].text;
    const enhancedPRD = extractJSON(aiResponse);

    res.json({
      success: true,
      data: {
        productName, initiativeType, featureFamily, release,
        regulatoryPermission, procurement, poLead, poMember, pmo,
        ...enhancedPRD
      },
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to generate PRD', message: error.message, details: error.response?.data || null });
  }
});

// ============================================
// CHAT ENDPOINT
// ============================================

app.post('/api/chat', async (req, res) => {
  try {
    const { message } = req.body;

    if (!message || message.trim().length === 0) {
      return res.status(400).json({ error: 'Message is required' });
    }

    const claudeResponse = await axios.post(
      'https://api.anthropic.com/v1/messages',
      {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 1024,
        system: CHAT_SYSTEM_PROMPT,
        messages: [{ role: 'user', content: message }]
      },
      {
        headers: {
          'x-api-key': CLAUDE_API_KEY,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        }
      }
    );

    const reply = claudeResponse.data.content[0].text;
    res.json({ success: true, reply, timestamp: new Date().toISOString() });

  } catch (error) {
    console.error('Chat error:', error.response?.data || error.message);
    res.status(500).json({ success: false, error: 'Chat failed', message: error.message });
  }
});

// ============================================
// EXPORT ENDPOINTS
// ============================================

app.post('/api/export-word', async (req, res) => {
  try {
    const {
      productName, initiativeType, featureFamily, release,
      regulatoryPermission, procurement, poLead, poMember, pmo,
      overview, problemStatement, targetUsers, keyFeatures, successMetrics, notes
    } = req.body;

    const sections = [
      new Paragraph({ text: 'PRODUCT REQUIREMENT DOCUMENT', heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
      new Paragraph({ text: productName || 'N/A', heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER, spacing: { after: 400 } }),
      new Paragraph({ text: 'PROJECT INFORMATION', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 200 } }),
      new Paragraph({ children: [new TextRun({ text: 'Product Name: ', bold: true }), new TextRun(productName || 'N/A')], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: 'Initiative Type: ', bold: true }), new TextRun(initiativeType || 'N/A')], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: 'Feature Family: ', bold: true }), new TextRun(featureFamily || 'N/A')], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: 'Release: ', bold: true }), new TextRun(release || 'N/A')], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: 'Regulatory Permission: ', bold: true }), new TextRun(regulatoryPermission || 'N/A')], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: 'Procurement: ', bold: true }), new TextRun(procurement || 'N/A')], spacing: { after: 400 } }),
      new Paragraph({ text: 'TEAM', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 200 } }),
      new Paragraph({ children: [new TextRun({ text: 'PO Lead: ', bold: true }), new TextRun(poLead || 'N/A')], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: 'PO Member: ', bold: true }), new TextRun(poMember || 'N/A')], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: 'PMO: ', bold: true }), new TextRun(pmo || 'N/A')], spacing: { after: 400 } }),
      new Paragraph({ children: [new PageBreak()] }),
    ];

    const addWordSection = (title, content) => {
      sections.push(new Paragraph({ text: title, heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 200 } }));
      const stripped = stripMarkdown(content || 'N/A');
      stripped.split('\n').forEach(line => {
        if (line.trim()) {
          sections.push(new Paragraph({ text: line.trim(), spacing: { line: 240, lineRule: 'auto', after: 100 } }));
        }
      });
    };

    addWordSection('OVERVIEW', overview);
    addWordSection('PROBLEM STATEMENT', problemStatement);
    addWordSection('TARGET USERS & PERSONAS', targetUsers);
    sections.push(new Paragraph({ children: [new PageBreak()] }));
    addWordSection('KEY FEATURES', keyFeatures);
    addWordSection('SUCCESS METRICS & KPIs', successMetrics);
    if (notes) addWordSection('ADDITIONAL NOTES', notes);

    sections.push(new Paragraph({ text: `Generated on ${new Date().toLocaleString()}`, alignment: AlignmentType.CENTER, spacing: { before: 400 } }));

    const doc = new Document({ sections: [{ children: sections }] });
    const buffer = await Packer.toBuffer(doc);
    res.contentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="PRD_${(productName || 'Document').replace(/\s+/g, '_')}_${Date.now()}.docx"`);
    res.send(buffer);

  } catch (error) {
    console.error('Word export error:', error.message);
    res.status(500).json({ error: 'Export to Word failed', message: error.message });
  }
});

app.post('/api/export-pdf', async (req, res) => {
  try {
    const {
      productName, initiativeType, featureFamily, release,
      regulatoryPermission, procurement, poLead, poMember, pmo,
      overview, problemStatement, targetUsers, keyFeatures, successMetrics, notes
    } = req.body;

    const doc = new PDFDocument({ bufferPages: true, margin: 50 });
    const chunks = [];

    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => {
      const pdf = Buffer.concat(chunks);
      res.contentType('application/pdf');
      res.send(pdf);
    });

    const addPDFSection = (title, content) => {
      doc.fontSize(14).font('Helvetica-Bold').text(title, { underline: true });
      doc.moveDown(0.3);
      const stripped = stripMarkdown(content || 'N/A');
      stripped.split('\n').forEach(line => {
        if (line.trim()) {
          doc.fontSize(11).font('Helvetica').text(line.trim(), { align: 'justify' });
        }
      });
      doc.moveDown(0.8);
    };

    doc.fontSize(20).font('Helvetica-Bold').text('Product Requirement Document', { align: 'center' });
    doc.fontSize(14).font('Helvetica').text(productName || 'N/A', { align: 'center' });
    doc.moveDown(1);

    doc.fontSize(11).font('Helvetica-Bold').text('Project Information', { underline: true });
    doc.fontSize(10).font('Helvetica');
    doc.text(`Product Name: ${productName || 'N/A'}`);
    doc.text(`Initiative Type: ${initiativeType || 'N/A'}`);
    doc.text(`Feature Family: ${featureFamily || 'N/A'}`);
    doc.text(`Release: ${release || 'N/A'}`);
    doc.text(`Regulatory Permission: ${regulatoryPermission || 'N/A'}`);
    doc.text(`Procurement: ${procurement || 'N/A'}`);
    doc.moveDown(0.5);

    doc.fontSize(11).font('Helvetica-Bold').text('Team', { underline: true });
    doc.fontSize(10).font('Helvetica');
    doc.text(`PO Lead: ${poLead || 'N/A'}`);
    doc.text(`PO Member: ${poMember || 'N/A'}`);
    doc.text(`PMO: ${pmo || 'N/A'}`);
    doc.moveDown(1);

    addPDFSection('Overview', overview);
    addPDFSection('Problem Statement', problemStatement);
    addPDFSection('Target Users & Personas', targetUsers);
    addPDFSection('Key Features', keyFeatures);
    addPDFSection('Success Metrics & KPIs', successMetrics);
    if (notes) addPDFSection('Additional Notes', notes);

    doc.fontSize(9).font('Helvetica').text(`Generated on ${new Date().toLocaleString()}`, { align: 'center' });
    doc.end();

  } catch (error) {
    console.error('PDF export error:', error.message);
    res.status(500).json({ error: 'Export to PDF failed', message: error.message });
  }
});

// ============================================
// START SERVER
// ============================================

app.listen(PORT, () => {
  console.log(`🚀 PRD Generator Backend running on http://localhost:${PORT}`);
  console.log(`✅ Claude API Key: ${CLAUDE_API_KEY ? 'Configured' : 'MISSING'}`);
  console.log('');
  console.log('📋 ENDPOINTS:');
  console.log('  POST /api/chat');
  console.log('  POST /api/enhance-mom');
  console.log('  POST /api/generate-prd');
  console.log('  POST /api/export-word');
  console.log('  POST /api/export-pdf');
  console.log('  GET /health');
});

module.exports = app;
