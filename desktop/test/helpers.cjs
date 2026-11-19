const http = require("node:http");

async function startServer(handler) {
  const server = http.createServer(handler);
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  return {
    server,
    url: `http://127.0.0.1:${address.port}/v1`,
    async close() {
      if (!server.listening) return;
      await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    }
  };
}

function sseResponse(response, { text = "fake model response", finishReason = "stop", usage = { prompt_tokens: 50, completion_tokens: 40, total_tokens: 90 } } = {}) {
  response.writeHead(200, { "Content-Type": "text/event-stream" });
  const events = [
    { id: "chatcmpl-test", object: "chat.completion.chunk", model: "mock-model", choices: [{ index: 0, delta: { role: "assistant", content: text }, finish_reason: null }] },
    { id: "chatcmpl-test", object: "chat.completion.chunk", model: "mock-model", choices: [{ index: 0, delta: {}, finish_reason: finishReason }] },
    { id: "chatcmpl-test", choices: [], usage }
  ];
  for (const event of events) response.write(`data: ${JSON.stringify(event)}\n\n`);
  response.end("data: [DONE]\n\n");
}

async function readJson(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

module.exports = { readJson, sseResponse, startServer };
