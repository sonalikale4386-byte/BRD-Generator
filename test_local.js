/**
 * test_local.js — Local test (no bot/emulator needed)
 *
 * Usage:
 *   node test_local.js
 */
require('dotenv').config();
const path                = require('path');
const { ClaudeService }   = require('./bot/services/claudeService');
const { generateExcel }   = require('./bot/brdCore');

// ── Edit project details here ─────────────────────────────────────────────────
const INPUT = {
  projectName: 'Invoice Automation System',
  userName:    'Lead Technical Consultant',
  extractedText: `
Meeting Summary – Requirements Workshop
Date: 09 Apr 2026
Attendees: CFO, IT Manager, Finance Team Lead, Procurement Head, Consultant

Key Discussion Points:
- Current invoice process is fully manual — takes 5-7 days from receipt to payment
- Finance team processes 2,000+ invoices per month
- 15% error rate due to manual data entry
- No visibility into invoice status for vendors or internal teams
- Integration needed with SAP ERP (FI module) for GL posting
- Mobile approval required for managers travelling
- Target: reduce processing time to same day, error rate below 1%
- Budget: approved for 6-month delivery
- Go-live must be before financial year end (31 Mar 2027)

Requirements identified:
- OCR-based invoice capture (email, scan, portal upload)
- Automated 3-level approval workflow (team lead > manager > CFO for amounts > 1L)
- SAP ERP integration for PO matching and GL posting
- Vendor self-service portal for invoice submission and status tracking
- Email and SMS notifications at each workflow stage
- Exception handling for mismatched POs
- Reporting dashboard: processing times, approval bottlenecks, payment status
- Full audit trail for compliance (7-year retention)
- Role-based access control
- Integration with bank payment systems for automated payment initiation
  `,
};
// ─────────────────────────────────────────────────────────────────────────────

async function main() {
  console.log('\n🚀 BRD Generator — Local Test');
  console.log('─'.repeat(50));
  console.log(`Project  : ${INPUT.projectName}`);
  console.log(`AI Mode  : ${process.env.CLAUDE_API_KEY ? '🟢 Claude API' : '🟡 Mock data'}`);
  console.log('─'.repeat(50));

  console.log('\n⏳ Step 1/2: Generating BRD content...');
  const claude  = new ClaudeService();
  const brdData = await claude.generateBRD(INPUT);

  console.log('⏳ Step 2/2: Building Excel document...');
  const outputDir = require('path').join(__dirname, 'output');
  const filePath  = await generateExcel(brdData, outputDir);

  const reqs = brdData.requirements || [];
  console.log('\n✅ BRD Generated Successfully!');
  console.log('─'.repeat(50));
  console.log(`📄 File          : ${filePath}`);
  console.log(`📋 Requirements  : ${reqs.length}`);
  console.log(`   🔴 High       : ${reqs.filter(r => r.priority === 'High').length}`);
  console.log(`   🟡 Medium     : ${reqs.filter(r => r.priority === 'Medium').length}`);
  console.log(`   🟢 Low        : ${reqs.filter(r => r.priority === 'Low').length}`);
  console.log(`👥 Stakeholders  : ${(brdData.key_stakeholders || []).length}`);
  console.log(`📦 Scope modules : ${(brdData.scope_summary || []).length}`);
  console.log('─'.repeat(50));
  console.log('\n➡️  Open the file in Excel to review.\n');
}

main().catch(err => { console.error('❌ Error:', err.message); process.exit(1); });
