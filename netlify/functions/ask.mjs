/* Ask-a-question endpoint for the invoice builder.
   The API key lives in a Netlify environment variable and never reaches the
   browser. A shared password is required as well, so the deployed function
   isn't an open proxy to the key for anyone who finds the URL.

   The browser sends REDACTED rows only — client and caregiver names are
   replaced with tokens (client#3, cg#7) before they leave the machine, and
   swapped back for real names in the browser when the answer is displayed.
   Nothing here ever sees a name, address, email or phone number. */

import Anthropic from "@anthropic-ai/sdk";

const SYSTEM = `You help an agency owner reconcile three monthly childcare exports while she builds an invoice for Care.com.

The data you are given is REDACTED: people appear as opaque tokens like client#3 and cg#7. Use those tokens verbatim in your answer — the app swaps them back to real names before she reads it. Never invent a real name.

How to answer:
- Lead with the answer, then the evidence. Cite job numbers and dates.
- Reconcile across the three sources rather than trusting one. A job can be Cancelled under one caregiver and Done under another — that is a reassignment, and the cancelled row should not be billed as worked time.
- A mileage claim filed by someone other than the caregiver on the job is a red flag worth calling out.
- Say plainly when the data cannot answer the question. Never guess a figure.
- Be brief. A few sentences and a short list beat an essay.`;

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { "content-type": "application/json" } });

export default async (req) => {
  if (req.method !== "POST") return json({ error: "Use POST." }, 405);

  const password = process.env.ASK_PASSWORD;
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey)   return json({ error: "ANTHROPIC_API_KEY isn't set on this site. Netlify → Site configuration → Environment variables." }, 500);
  if (!password) return json({ error: "ASK_PASSWORD isn't set on this site. Netlify → Site configuration → Environment variables. Without it the function would be open to anyone." }, 500);
  if (req.headers.get("x-ask-password") !== password) return json({ error: "Wrong password." }, 401);

  let body;
  try { body = await req.json(); } catch { return json({ error: "Malformed request." }, 400); }
  const question = String(body?.question ?? "").trim();
  const data = body?.data;
  if (!question) return json({ error: "No question." }, 400);
  if (!data)     return json({ error: "No data — read the exports first." }, 400);

  const client = new Anthropic({ apiKey });
  const stream = client.messages.stream({
    model: "claude-opus-5",
    max_tokens: 8000,
    thinking: { type: "adaptive" },
    output_config: { effort: "medium" },   // a lookup over a few hundred rows, and latency matters
    system: SYSTEM,
    messages: [{
      role: "user",
      content: `Here is this month's redacted invoice data as JSON.\n\n${JSON.stringify(data)}\n\nQuestion: ${question}`,
    }],
  });

  // Stream the answer back so a long reply can't hit the function's time limit.
  const out = new ReadableStream({
    async start(controller) {
      const enc = new TextEncoder();
      try {
        for await (const event of stream) {
          if (event.type === "content_block_delta" && event.delta.type === "text_delta")
            controller.enqueue(enc.encode(event.delta.text));
        }
        const final = await stream.finalMessage();
        if (final.stop_reason === "refusal")
          controller.enqueue(enc.encode("\n\n[Claude declined to answer this one. Rephrasing usually helps.]"));
      } catch (err) {
        controller.enqueue(enc.encode(`\n\n[Couldn't finish: ${err?.message || err}]`));
      }
      controller.close();
    },
  });
  return new Response(out, { headers: { "content-type": "text/plain; charset=utf-8", "cache-control": "no-store" } });
};
