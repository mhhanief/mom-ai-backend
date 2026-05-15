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

// System Prompt untuk MoM Agent
const MOM_SYSTEM_PROMPT = `Anda adalah Professional Meeting Minutes (MoM) Agent dengan expertise tinggi dalam corporate documentation dan business communication.

PERAN & TANGGUNG JAWAB:
- Mengubah input meeting notes mentah menjadi professional Meeting Minutes format resmi
- Memastikan struktur, grammar, dan tone konsisten dengan standar corporate
- Menambahkan clarity, menghilangkan redundansi, dan menyusun ulang untuk readability
- Memastikan action items jelas dengan PIC dan deadline

GUIDELINES:
1. TONE: Professional, formal, clear, concise
2. STRUCTURE: Terstruktur rapi dengan section yang jelas
3. GRAMMAR: Gunakan Bahasa Indonesia formal atau English sesuai input
4. ACTION ITEMS: Format jelas - "[PIC Name] - [Action] - [Deadline]"
5. CLARITY: Ganti jargon yang tidak jelas dengan penjelasan singkat

OUTPUT FORMAT yang harus diikuti:
- Judul: Tetap gunakan judul original, tambahkan nomor/versi jika perlu
- Tanggal & Lokasi: Format konsisten
- Partisipan: List nama yang attend, highlight PIC jika ada
- Agenda: Bullet points singkat dan jelas
- Pembahasan: Expand dengan konteks, insight, dan context dari diskusi
- Kesepakatan: Ringkas kesimpulan dan keputusan yang diambil
- Action Items: List detail dengan owner & deadline yang jelas

JANGAN:
- Tambahkan informasi yang tidak ada di input
- Ubah keputusan/kesepakatan yang sudah ditetapkan
- Hilangkan detail penting
- Ubah struktur fundamental input

LAKUKAN:
- Polish grammar & spelling
- Improve clarity tanpa menambah info baru
- Standardize format & terminology
- Ensure consistency across sections`;

// System Prompt untuk PRD Agent
const PRD_SYSTEM_PROMPT = `Anda adalah Expert Product Requirement Document (PRD) Agent dengan deep expertise dalam product management, requirements definition, dan technical documentation.

PERAN & TANGGUNG JAWAB:
- Transform input PRD data menjadi comprehensive, well-structured PRD document
- Ensure semua sections clear, detailed, dan aligned dengan BNI RDC standards
- Maintain professional tone, consistent formatting, dan logical flow
- Identify gaps dan provide thoughtful suggestions tanpa mengubah core intent

GUIDELINES:
1. TONE: Professional, strategic, clear, stakeholder-ready
2. STRUCTURE: Logical flow dari Overview → Problem → Solution → Success Metrics
3. DETAIL LEVEL: Sufficient untuk development team guidance tanpa over-specification
4. LANGUAGE: Bahasa Indonesia formal atau English sesuai context
5. COMPLETENESS: Ensure semua sections well-developed dengan examples/details

OUTPUT REQUIREMENTS:
- Overview: Executive summary yang compelling
- Problem Statement: Jelas identify pain points dan business drivers
- Target Users: Specific personas dengan use cases
- Key Features: Detailed feature list dengan priorities/phases
- Success Metrics: Measurable KPIs untuk track product success
- Additional Notes: Dependencies, constraints, atau implementation considerations

JANGAN:
- Invent features yang bukan di input
- Change core requirements/decisions
- Lose important details dalam condensing

LAKUKAN:
- Enhance clarity dan depth
- Add structure dan organization
- Ensure completeness across all sections
- Provide polished, presentation-ready output`;

// System Prompt untuk Chat Assistant
const CHAT_SYSTEM_PROMPT = `Anda adalah AI Assistant bernama Archi yang membantu tim RDC BNI dalam membuat dokumen PRD dan Meeting Minutes dengan lebih cepat.

KARAKTERISTIK:
- Helpful, friendly, dan professional
- Gunakan Bahasa Indonesia informal tapi tetap profesional
- Fokus pada topik: PRD creation, Meeting Minutes, document best practices
- Jangan: berikan saran teknis di luar scope, klaim expertise yang tidak ada

CAPABILITIES:
- Discuss PRD structure dan best practices
- Suggest improvements untuk document quality
- Answer questions tentang fitur/workflow di PRD generator
- Help brainstorm ideas untuk feature/product definition
- Explain document-related concepts

TONE: Friendly, conversational, helpful. Use "aku" dan "kamu" untuk lebih personal.`;

// Route health check
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// ============================================
// MEETING MINUTES ENDPOINTS
// ============================================

// Route enhance MoM dengan Claude
app.post('/api/enhance-mom', async (req, res) => {
  try {
    const { judul, tanggal, lokasi, partisipan, agenda, pembahasan, kesepakatan, actionItems } = req.body;

    // Validasi input
    if (!judul || !tanggal || !lokasi || !partisipan || !agenda || !pembahasan || !kesepakatan || !actionItems) {
      return res.status(400).json({
        error: 'Missing required fields',
        required: ['judul', 'tanggal', 'lokasi', 'partisipan', 'agenda', 'pembahasan', 'kesepakatan', 'actionItems']
      });
    }

    // Format input untuk AI
    const userPrompt = `Berikut adalah raw input Meeting Minutes yang perlu di-enhance dan di-polish menjadi professional format:

JUDUL MEETING: ${judul}
TANGGAL: ${tanggal}
LOKASI: ${lokasi}
PARTISIPAN: ${partisipan}

AGENDA MEETING:
${agenda}

PEMBAHASAN MEETING:
${pembahasan}

KESEPAKATAN MEETING:
${kesepakatan}

ACTION ITEMS:
${actionItems}

---

Mohon enhance dan polish semua section di atas menjadi professional Meeting Minutes format, mengikuti guidelines yang telah diberikan.

Output harus HANYA dalam format JSON (tanpa text lain):
{
  "judul": "string",
  "tanggal": "string (format: DD Bulan YYYY)",
  "lokasi": "string",
  "partisipan": "string",
  "agenda": "string",
  "pembahasan": "string",
  "kesepakatan": "string",
  "actionItems": "string"
}`;

    // Call Claude API
    const claudeResponse = await axios.post(
      'https://api.anthropic.com/v1/messages',
      {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 2048,
        system: MOM_SYSTEM_PROMPT,
        messages: [
          {
            role: 'user',
            content: userPrompt
          }
        ]
      },
      {
        headers: {
          'x-api-key': CLAUDE_API_KEY,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        }
      }
    );

    // Parse response
    const aiResponse = claudeResponse.data.content[0].text;
    
    // Extract JSON dari response
    let enhancedMoM;
    try {
      enhancedMoM = JSON.parse(aiResponse);
    } catch (e) {
      // Coba remove markdown code block wrapper
      let cleanedResponse = aiResponse
        .replace(/```json\n?/g, '')
        .replace(/```\n?/g, '')
        .trim();
      
      try {
        enhancedMoM = JSON.parse(cleanedResponse);
      } catch (e2) {
        // Last resort: try extract JSON object
        const jsonMatch = cleanedResponse.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          try {
            enhancedMoM = JSON.parse(jsonMatch[0]);
          } catch (e3) {
            throw new Error('Could not parse AI response as JSON: ' + cleanedResponse.substring(0, 200));
          }
        } else {
          throw new Error('Could not extract JSON from AI response');
        }
      }
    }

    // Return enhanced MoM
    res.json({
      success: true,
      data: enhancedMoM,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({
      error: 'Failed to enhance MoM',
      message: error.message,
      details: error.response?.data || null
    });
  }
});

// ============================================
// PRD ENDPOINTS
// ============================================

// Route generate PRD dengan Claude AI
app.post('/api/generate-prd', async (req, res) => {
  try {
    const {
      productName,
      initiativeType,
      featureFamily,
      release,
      regulatoryPermission,
      procurement,
      poLead,
      poMember,
      pmo,
      overview,
      problemStatement,
      targetUsers,
      keyFeatures,
      successMetrics,
      notes
    } = req.body;

    // Validasi minimal required fields
    if (!productName || !overview || !problemStatement || !targetUsers || !keyFeatures || !successMetrics) {
      return res.status(400).json({
        error: 'Missing required fields for PRD generation'
      });
    }

    // Format input untuk AI
    const userPrompt = `Berikut adalah PRD input yang perlu di-enhance dan di-develop menjadi comprehensive PRD document:

PRODUCT NAME: ${productName}
INITIATIVE TYPE: ${initiativeType}
FEATURE FAMILY: ${featureFamily}
RELEASE: ${release}
REGULATORY PERMISSION: ${regulatoryPermission}
PROCUREMENT: ${procurement}

TEAM:
- PO Lead: ${poLead}
- PO Member: ${poMember}
- PMO: ${pmo}

---

OVERVIEW:
${overview}

PROBLEM STATEMENT:
${problemStatement}

TARGET USERS/PERSONAS:
${targetUsers}

KEY FEATURES:
${keyFeatures}

SUCCESS METRICS/KPI:
${successMetrics}

ADDITIONAL NOTES:
${notes || 'None'}

---

Mohon enhance dan develop semua section di atas menjadi comprehensive, professional PRD document dengan level of detail yang appropriate untuk development team guidance.

Output harus HANYA dalam format JSON (tanpa text lain):
{
  "overview": "string (enhanced overview dengan lebih detail)",
  "problemStatement": "string (detailed problem statement dengan business context)",
  "targetUsers": "string (detailed personas dengan use cases)",
  "keyFeatures": "string (comprehensive feature list dengan details)",
  "successMetrics": "string (detailed KPIs dengan measurement approach)"
}`;

    // Call Claude API
    const claudeResponse = await axios.post(
      'https://api.anthropic.com/v1/messages',
      {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 3000,
        system: PRD_SYSTEM_PROMPT,
        messages: [
          {
            role: 'user',
            content: userPrompt
          }
        ]
      },
      {
        headers: {
          'x-api-key': CLAUDE_API_KEY,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        }
      }
    );

    // Parse response
    const aiResponse = claudeResponse.data.content[0].text;
    
    // Extract JSON dari response
    let enhancedPRD;
    try {
      enhancedPRD = JSON.parse(aiResponse);
    } catch (e) {
      // Coba remove markdown code block wrapper
      let cleanedResponse = aiResponse
        .replace(/```json\n?/g, '')
        .replace(/```\n?/g, '')
        .trim();
      
      try {
        enhancedPRD = JSON.parse(cleanedResponse);
      } catch (e2) {
        // Last resort: try extract JSON object
        const jsonMatch = cleanedResponse.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          try {
            enhancedPRD = JSON.parse(jsonMatch[0]);
          } catch (e3) {
            throw new Error('Could not parse AI response as JSON: ' + cleanedResponse.substring(0, 200));
          }
        } else {
          throw new Error('Could not extract JSON from AI response');
        }
      }
    }

    // Return enhanced PRD dengan original fields
    res.json({
      success: true,
      data: {
        productName,
        initiativeType,
        featureFamily,
        release,
        regulatoryPermission,
        procurement,
        poLead,
        poMember,
        pmo,
        ...enhancedPRD
      },
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({
      error: 'Failed to generate PRD',
      message: error.message,
      details: error.response?.data || null
    });
  }
});

// ============================================
// CHAT ENDPOINT
// ============================================

// Route chat dengan Claude
app.post('/api/chat', async (req, res) => {
  try {
    const { message } = req.body;

    if (!message || message.trim().length === 0) {
      return res.status(400).json({
        error: 'Message is required'
      });
    }

    // Call Claude API
    const claudeResponse = await axios.post(
      'https://api.anthropic.com/v1/messages',
      {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 1024,
        system: CHAT_SYSTEM_PROMPT,
        messages: [
          {
            role: 'user',
            content: message
          }
        ]
      },
      {
        headers: {
          'x-api-key': CLAUDE_API_KEY,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        }
      }
    );

    // Extract reply
    const reply = claudeResponse.data.content[0].text;

    res.json({
      success: true,
      reply: reply,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('Chat error:', error.response?.data || error.message);
    res.status(500).json({
      success: false,
      error: 'Chat failed',
      message: error.message
    });
  }
});

// ============================================
// EXPORT ENDPOINTS
// ============================================

// Route export to Word
app.post('/api/export-word', async (req, res) => {
  try {
    const {
      productName,
      initiativeType,
      featureFamily,
      release,
      regulatoryPermission,
      procurement,
      poLead,
      poMember,
      pmo,
      overview,
      problemStatement,
      targetUsers,
      keyFeatures,
      successMetrics,
      notes
    } = req.body;

    const sections = [
      new Paragraph({
        text: 'PRODUCT REQUIREMENT DOCUMENT',
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 }
      }),
      new Paragraph({
        text: productName || 'N/A',
        heading: HeadingLevel.HEADING_2,
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 }
      }),

      new Paragraph({
        text: 'PROJECT INFORMATION',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Product Name: ', bold: true }),
          new TextRun(productName || 'N/A')
        ],
        spacing: { after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Initiative Type: ', bold: true }),
          new TextRun(initiativeType || 'N/A')
        ],
        spacing: { after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Feature Family: ', bold: true }),
          new TextRun(featureFamily || 'N/A')
        ],
        spacing: { after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Release: ', bold: true }),
          new TextRun(release || 'N/A')
        ],
        spacing: { after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Regulatory Permission: ', bold: true }),
          new TextRun(regulatoryPermission || 'N/A')
        ],
        spacing: { after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Procurement: ', bold: true }),
          new TextRun(procurement || 'N/A')
        ],
        spacing: { after: 400 }
      }),

      new Paragraph({
        text: 'TEAM',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'PO Lead: ', bold: true }),
          new TextRun(poLead || 'N/A')
        ],
        spacing: { after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'PO Member: ', bold: true }),
          new TextRun(poMember || 'N/A')
        ],
        spacing: { after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'PMO: ', bold: true }),
          new TextRun(pmo || 'N/A')
        ],
        spacing: { after: 400 }
      }),

      new Paragraph({ children: [new PageBreak()] }),

      new Paragraph({
        text: 'OVERVIEW',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 }
      }),
      new Paragraph({
        text: overview || 'N/A',
        spacing: { line: 240, lineRule: 'auto', after: 200 }
      }),

      new Paragraph({
        text: 'PROBLEM STATEMENT',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 }
      }),
      new Paragraph({
        text: problemStatement || 'N/A',
        spacing: { line: 240, lineRule: 'auto', after: 200 }
      }),

      new Paragraph({
        text: 'TARGET USERS & PERSONAS',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 }
      }),
      new Paragraph({
        text: targetUsers || 'N/A',
        spacing: { line: 240, lineRule: 'auto', after: 200 }
      }),

      new Paragraph({ children: [new PageBreak()] }),

      new Paragraph({
        text: 'KEY FEATURES',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 }
      }),
      new Paragraph({
        text: keyFeatures || 'N/A',
        spacing: { line: 240, lineRule: 'auto', after: 200 }
      }),

      new Paragraph({
        text: 'SUCCESS METRICS & KPIs',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 }
      }),
      new Paragraph({
        text: successMetrics || 'N/A',
        spacing: { line: 240, lineRule: 'auto', after: 400 }
      })
    ];

    if (notes) {
      sections.push(
        new Paragraph({
          text: 'ADDITIONAL NOTES',
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 200, after: 200 }
        }),
        new Paragraph({
          text: notes,
          spacing: { line: 240, lineRule: 'auto', after: 400 }
        })
      );
    }

    sections.push(
      new Paragraph({
        text: `Generated on ${new Date().toLocaleString()}`,
        alignment: AlignmentType.CENTER,
        spacing: { before: 400 }
      })
    );

    const doc = new Document({
      sections: [{ children: sections }]
    });

    const buffer = await Packer.toBuffer(doc);
    res.contentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="PRD_${(productName || 'Document').replace(/\s+/g, '_')}_${Date.now()}.docx"`);
    res.send(buffer);

  } catch (error) {
    console.error('Word export error:', error.message);
    res.status(500).json({
      error: 'Export to Word failed',
      message: error.message
    });
  }
});

// Route export to PDF
app.post('/api/export-pdf', async (req, res) => {
  try {
    const {
      productName,
      initiativeType,
      featureFamily,
      release,
      regulatoryPermission,
      procurement,
      poLead,
      poMember,
      pmo,
      overview,
      problemStatement,
      targetUsers,
      keyFeatures,
      successMetrics,
      notes
    } = req.body;

    // Create PDF document
    const doc = new PDFDocument({ bufferPages: true, margin: 50 });
    const chunks = [];

    // Collect chunks for response
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => {
      const pdf = Buffer.concat(chunks);
      res.contentType('application/pdf');
      res.send(pdf);
    });

    // Helper function untuk add section
    const addSection = (title, content) => {
      doc.fontSize(14).font('Helvetica-Bold').text(title, { underline: true });
      doc.fontSize(11).font('Helvetica').text(content, { align: 'justify' });
      doc.moveDown(0.5);
    };

    // Title
    doc.fontSize(20).font('Helvetica-Bold').text('Product Requirement Document', { align: 'center' });
    doc.fontSize(14).font('Helvetica').text(productName || 'N/A', { align: 'center' });
    doc.moveDown(1);

    // Metadata
    doc.fontSize(11).font('Helvetica-Bold').text('Project Information', { underline: true });
    doc.fontSize(10).font('Helvetica');
    doc.text(`Product Name: ${productName || 'N/A'}`);
    doc.text(`Initiative Type: ${initiativeType || 'N/A'}`);
    doc.text(`Feature Family: ${featureFamily || 'N/A'}`);
    doc.text(`Release: ${release || 'N/A'}`);
    doc.text(`Regulatory Permission: ${regulatoryPermission || 'N/A'}`);
    doc.text(`Procurement: ${procurement || 'N/A'}`);
    doc.moveDown(0.5);

    // Team
    doc.fontSize(11).font('Helvetica-Bold').text('Team', { underline: true });
    doc.fontSize(10).font('Helvetica');
    doc.text(`PO Lead: ${poLead || 'N/A'}`);
    doc.text(`PO Member: ${poMember || 'N/A'}`);
    doc.text(`PMO: ${pmo || 'N/A'}`);
    doc.moveDown(1);

    // Overview
    addSection('Overview', overview || 'N/A');

    // Problem Statement
    addSection('Problem Statement', problemStatement || 'N/A');

    // Target Users
    addSection('Target Users & Personas', targetUsers || 'N/A');

    // Key Features
    addSection('Key Features', keyFeatures || 'N/A');

    // Success Metrics
    addSection('Success Metrics & KPIs', successMetrics || 'N/A');

    // Additional Notes
    if (notes) {
      addSection('Additional Notes', notes);
    }

    // Footer with timestamp
    doc.fontSize(9).font('Helvetica').text(`Generated on ${new Date().toLocaleString()}`, { align: 'center', color: '#999' });

    // End PDF
    doc.end();

  } catch (error) {
    console.error('PDF export error:', error.message);
    res.status(500).json({
      error: 'Export to PDF failed',
      message: error.message
    });
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
  console.log('  POST /api/chat - Chat dengan AI Assistant');
  console.log('  POST /api/enhance-mom - Enhance Meeting Minutes dengan Claude AI');
  console.log('  POST /api/generate-prd - Generate PRD dengan Claude AI (toggle ON)');
  console.log('  POST /api/export-word - Export PRD ke Word');
  console.log('  POST /api/export-pdf - Export PRD ke PDF');
  console.log('  GET /health - Health check');
  console.log('');
});

module.exports = app;
