"""Smoke test the installed DSH process and its streaming company gateway route."""
import asyncio
import json
import os
import threading
from pathlib import Path
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

PLAN = {"title":"DSH 验收", "subtitle":"网关协议测试",
        "sections":[{"heading":"进度","points":["已走通实际 SDK 调用。"]}]}
MODEL = "deepseek-v4-flash"
CAPTURED = []


class FakeGateway(BaseHTTPRequestHandler):
    def do_POST(self):
        data = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        CAPTURED.append((self.path, dict(self.headers), data))
        self.send_response(200)
        self.send_header("Content-Type","text/event-stream")
        self.end_headers()
        tools={tool["function"]["name"] for tool in data.get("tools",[])}
        tool_result=next((message["content"] for message in data["messages"] if message["role"]=="tool"),None)
        if tools and tool_result is None:
            name="office_context_read" if "office_context_read" in tools else "todo_write"
            arguments={} if name=="office_context_read" else {"todos":[{"content":"核对会议纪要","status":"completed"}]}
            answer_delta={"role":"assistant","tool_calls":[{"index":0,"id":"call_acceptance","type":"function",
                          "function":{"name":name,"arguments":json.dumps(arguments,ensure_ascii=False)}}]}
            end="tool_calls"
        else:
            answer=({"tool_result":tool_result} if tool_result is not None else
                    {"action":"export","reply":"方案已整理","kind":"docx","plan":PLAN} if "办公助理" in data["messages"][0]["content"] else PLAN)
            answer_delta={"role":"assistant","content":json.dumps(answer,ensure_ascii=False)}
            end="stop"
        for delta, status in [
            (answer_delta,None),
            ({},end),
        ]:
            chunk = {"id":"chatcmpl-acceptance","object":"chat.completion.chunk","created":1,"model":MODEL,
                     "choices":[{"index":0,"delta":delta,"finish_reason":status}]}
            self.wfile.write(("data: " + json.dumps(chunk) + "\n\n").encode())
        self.wfile.write(b'data: {"id":"chatcmpl-acceptance","choices":[],"usage":{"prompt_tokens":50,"completion_tokens":40,"total_tokens":90}}\n\ndata: [DONE]\n\n')

    def log_message(self, *_):
        return


def main():
    with ThreadingHTTPServer(("127.0.0.1",0), FakeGateway) as stub:
        thread = threading.Thread(target=stub.serve_forever, daemon=True)
        thread.start()
        try:
            os.environ.update(QO_DEMO="1", QO_ENGINE="dsh", QO_GATEWAY_MODEL=MODEL, QO_GATEWAY_KEY="only-for-test",
                              QO_GATEWAY_BASE_URL=f"http://127.0.0.1:{stub.server_port}/v1")
            from server.gateway import plan_for, validate
            validate()
            plan, usage = asyncio.run(plan_for("docx","验收","写关于公司协作的内部文书",None,[],5,"lin","test-run"))
            assert plan.model_dump() == PLAN
            assert usage == {"prompt_tokens":50,"completion_tokens":40,"total_tokens":90}
            assert len(CAPTURED) == 1
            path, headers, data = CAPTURED[0]
            assert path == "/v1/chat/completions"
            assert headers.get("X-QO-Run") == "test-run" and headers.get("X-QO-Employee") == "6c696e"
            assert headers.get("Authorization") == "Bearer only-for-test"
            assert data["model"] == MODEL and data["stream"] and not data.get("tools")
            from server.conversations import TurnRequest, decide, office_file
            body = TurnRequest(id="4bc7d684-a491-401e-9242-6423073245ca",message="把上轮方案导出 Word")
            earlier = [{"request":"整理协作方案","reply":"已经梳理关键步骤"}]
            decision, usage = asyncio.run(decide(body,None,[],earlier,None,{"id":"lin"},"next-run"))
            kind, _, filename, raw = office_file(decision,body.mode)
            assert kind == "docx" and filename.endswith(".docx") and raw.startswith(b"PK")
            assert usage["total_tokens"] == 90 and len(CAPTURED) == 2
            assert "整理协作方案" in str(CAPTURED[1][2]["messages"])
            from server import gateway
            response,_=asyncio.run(gateway.request_model("调用 office_context_read","native-context",
                gateway.headers_for("lin",None,"native-context"),"先查看授权的项目资料",{"资料":"原生工具读取证明"}))
            assert "原生工具读取证明" in response and len(CAPTURED)==4
            gateway.EXTRA_PATCH=str(Path(__file__).resolve().parents[1]/"harness/todo.example.patch.yml")
            gateway.EXTRA_TOOLS=frozenset({"todo_write"})
            gateway.validate()
            response,_=asyncio.run(gateway.request_model("用 todo_write 记录已完成项","native-upstream",
                gateway.headers_for("lin",None,"native-upstream"),"使用计划工具"))
            assert "1 completed" in response and len(CAPTURED)==6,(response,len(CAPTURED))
            print("DSH Word export, native Office context, and unchanged upstream todo_write passed.")
        finally:
            stub.shutdown()
            thread.join()


if __name__ == "__main__":
    main()
