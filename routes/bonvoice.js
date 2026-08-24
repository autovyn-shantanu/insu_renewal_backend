const axios = require("axios");
const crypto = require("crypto");

// ===================== CONFIG =====================
const BONVOICE_CONFIG = {
  BASE_URL: process.env.BONVOICE_BASE_URL || "https://voiceai.bonvoice.com/api/v1",
  PUBLIC_BASE_URL:
  process.env.BONVOICE_PUBLIC_BASE_URL || "https://voiceai.bonvoice.com/api/public",
  TOKEN: process.env.BONVOICE_TOKEN || "cm_live_d082244083cd123e48c32b228bdaab7334aade3e8341b785",
};

if (!BONVOICE_CONFIG.TOKEN) {
  console.warn("[Bonvoice] BONVOICE_TOKEN missing. Set process.env.BONVOICE_TOKEN");
}

const client = axios.create({
  baseURL: BONVOICE_CONFIG.BASE_URL,
  headers: {
    "Content-Type": "application/json",
    Authorization: `Bearer ${BONVOICE_CONFIG.TOKEN}`,
  },
  timeout: 30000,
});

function toBonvoiceError(err, fallbackMsg) {
  const data = err?.response?.data;
  const status = err?.response?.status;
  const msg = data?.message || err?.message || fallbackMsg;
  const e = new Error(msg);
  e.status = status;
  e.data = data;
  return e;
}

// ===================== LEAD =====================

/**
 * POST /api/v1/leads
 * Creates a lead. If payload.triggerCall=true then call is initiated immediately.
 * ✅ Bonvoice automatically sends webhook when call completes
 */
async function createLead(payload = {}) {
  try {
    const res = await client.post("/leads", payload);
    return res.data;
  } catch (err) {
    throw toBonvoiceError(err, "Bonvoice createLead failed");
  }
}

/**
 * ✅ MAIN FUNCTION: Create lead + trigger call
 * Bonvoice automatically sends webhook after call completes
 * (callId usually null in response; webhook me milega)
 */
async function createLeadAndCall({
  name,
  phone,
  email,
  college,
  company,
  program,
  column1,      // ✅ Custom field 1
  column2,      // ✅ Custom field 2
  column3, 
  column7,     // ✅ Custom field 3
  promptName,   // Bonvoice prompt name
} = {}) {
  if (!phone) throw new Error("phone is required");

  console.log("[BONVOICE API] Creating lead with call trigger:", {
    name,
    phone,
    program,
    column1,
    column2,
    column3,
    promptName,
  });

  // ✅ triggerCall: true = Bonvoice automatically करेगा call initiate
  // ✅ Webhook automatically आएगा जब call complete हो
  return createLead({
    name: name ?? null,
    phone,
    email: email ?? null,
    college: college ?? null,
    company: company ?? null,
    program: program ?? null,
    column1: column1 ?? null,      // ✅ Custom data
    column2: column2 ?? null,      // ✅ Custom data
    column3: column3 ?? null,
    column7:column7 ?? null,      // ✅ Custom data
    triggerCall: true,             // ✅ Automatically trigger call
    ...(promptName ? { promptName } : {}),
  });
}

// ===================== CALL =====================

/**
 * POST /api/v1/leads/:leadId/call
 * Trigger call for existing leadId (manual)
 * Webhook आएगा जब call complete हो
 */
async function triggerSingleCall(leadId, options = {}) {
  if (!leadId) throw new Error("leadId is required");

  const body = {};
  if (options.promptName) body.promptName = options.promptName;
  if (options.program !== undefined) body.program = options.program;

  try {
    console.log("[BONVOICE API] Triggering call for leadId:", leadId);
    const res = await client.post(`/leads/${leadId}/call`, body);
    return res.data;
  } catch (err) {
    throw toBonvoiceError(err, "Bonvoice triggerSingleCall failed");
  }
}

/**
 * 2-step helper (agar callId immediately chahiye):
 * create lead (triggerCall=false) -> trigger call manually
 */
async function createLeadThenTriggerCall(leadPayload = {}, callOptions = {}) {
  const lead = await createLead({ ...leadPayload, triggerCall: false });
  const leadId = lead?.leadId;
  if (!leadId) throw new Error("Lead created but leadId missing in response");

  const call = await triggerSingleCall(leadId, callOptions);
  return {
    leadId,
    callId: call?.callId || null,
    lead,
    call,
  };
}

/**
 * ✅ Batch calls - Bonvoice automatically भेजेगा webhook for each call
 * Input format:
 * [
 *   { leadPayload: {...}, callOptions: {...} },
 *   ...
 * ]
 */
async function triggerBatchCalls(leads = [], { concurrency = 10 } = {}) {
  if (!Array.isArray(leads) || !leads.length) {
    throw new Error("leads array must have at least 1 entry");
  }
  if (leads.length > 200) {
    throw new Error("leads array must have <= 200 entries");
  }

  console.log("[BONVOICE API] Triggering batch calls:", leads.length);

  const results = new Array(leads.length);
  let idx = 0;

  const worker = async () => {
    while (idx < leads.length) {
      const current = idx++;
      const item = leads[current];

      try {
        const leadData = await createLeadAndCall(item);
        results[current] = {
          ok: true,
          leadId: leadData?.leadId,
          phone: item.phone,
          data: leadData,
        };
        console.log(`[BONVOICE API] Lead ${current + 1} created`, {
          leadId: leadData?.leadId,
          phone: item.phone,
        });
      } catch (e) {
        results[current] = {
          ok: false,
          phone: item.phone,
          error: e.message,
          status: e.status,
        };
        console.error(`[BONVOICE API] Lead ${current + 1} failed:`, e.message);
      }
    }
  };

  const workers = Array.from(
    { length: Math.min(concurrency, leads.length) },
    worker
  );

  await Promise.all(workers);
  return results;
}

// ===================== READ =====================

/**
 * GET /api/v1/leads/:leadId
 * Returns lead + latestCall (transcript/summary/recording_url etc)
 */
async function getLead(leadId) {
  if (!leadId) throw new Error("leadId is required");
  try {
    const res = await client.get(`/leads/${leadId}`);
    return res.data;
  } catch (err) {
    throw toBonvoiceError(err, "Bonvoice getLead failed");
  }
}

/**
 * Get latest call status using lead data
 */
async function getLatestCallStatus(leadId) {
  const lead = await getLead(leadId);
  return {
    leadId: lead.id,
    leadStatus: lead.leadStatus,
    conversionScore: lead.conversionScore,
    latestCall: lead.latestCall || null,
  };
}

/**
 * GET /api/public/recordings/:callId (public, no auth)
 * Streams MP3 to Express response
 */
async function getCallRecording(callId, res) {
  if (!callId) throw new Error("callId is required");

  try {
    const response = await axios.get(
      `${BONVOICE_CONFIG.PUBLIC_BASE_URL}/recordings/${callId}`,
      { responseType: "stream", timeout: 30000 }
    );

    res.setHeader(
      "Content-Type",
      response.headers["content-type"] || "audio/mpeg"
    );
    response.data.pipe(res);
  } catch (err) {
    throw toBonvoiceError(err, "Bonvoice recording download failed");
  }
}

// ===================== WEBHOOK SIGNATURE VERIFY =====================

/**
 * Bonvoice webhook signature verify (HMAC-SHA256)
 * signature header: X-Bonvoice-Signature: sha256=<hex>
 *
 * IMPORTANT: rawBody must be Buffer (express.raw())
 */
function verifyBonvoiceWebhook(rawBody, signatureHeader, secret) {
  if (!rawBody || !signatureHeader || !secret) return false;

  // rawBody must be Buffer (express.raw middleware)
  const sig = String(signatureHeader).trim(); // e.g. "sha256=abcd123..."

  const expectedHex = crypto
    .createHmac("sha256", secret) // ✅ algorithm + secret
    .update(rawBody)              // ✅ Buffer
    .digest("hex");

  const expected = `sha256=${expectedHex}`; // ✅ correct prefix

  const a = Buffer.from(sig, "utf8");
  const b = Buffer.from(expected, "utf8");
  if (a.length !== b.length) return false;

  return crypto.timingSafeEqual(a, b);
}

// ===================== EXPORTS =====================
module.exports = {
  // lead
  createLead,
  createLeadAndCall,           // ✅ MAIN - Auto trigger call + webhook
  createLeadThenTriggerCall,   // Manual 2-step

  // call
  triggerSingleCall,
  triggerBatchCalls,           // ✅ Batch with auto webhook

  // read
  getLead,
  getLatestCallStatus,
  getCallRecording,

  // webhook
  verifyBonvoiceWebhook,
};