import itertools
import re
import unittest

from server import pipeline, prompting
from server.presets import ALL, BY_ID, GENRES, OPERATIONS

# (请求, 期望的操作规范 id 或 None, 期望的文体规范 id 或 None)
ROUTES = [
    ("帮我把今天的会议笔记整理成纪要", None, "minutes"),
    ("整理一下今天例会的录音转写", None, "minutes"),
    ("写一份本周周报", None, "periodic"),
    ("帮我写个年终总结", None, "summary"),
    ("总结一下这篇文章", "summarize", None),
    ("做一份下季度工作计划PPT，10页", None, "plan"),
    ("发个放假通知", None, "notice"),
    ("写一份会议通知", None, "notice"),
    ("起草一份关于申请购置办公设备的请示", None, "official"),
    ("写一封邀请函，邀请王总参加发布会", None, "letter"),
    ("给客户写封邮件跟进报价", None, "email"),
    ("写封邮件通知大家下周停电", None, "email"),
    ("写一篇公众号推文介绍我们的新产品", None, "press"),
    ("写一个发布会新闻稿", None, "press"),
    ("帮我写一个年会主持词", None, "speech"),
    ("策划一个30人的团建活动", None, "event"),
    ("写一份行业分析报告", None, "research"),
    ("帮我写高新技术企业申报书", None, "proposal"),
    ("写一份项目立项报告", None, "proposal"),
    ("起草一份报销管理制度", None, "policy"),
    ("帮我审一下这份采购合同有没有风险", None, "contract_review"),
    ("帮我看看这份劳动合同", None, "contract_review"),
    ("起草一份技术服务合同", None, "contract_draft"),
    ("写一个Java开发的招聘JD", None, "hr_recruit"),
    ("写一份录用通知", None, "hr_recruit"),
    ("写一段绩效评语", None, "hr_people"),
    ("解除劳动合同通知书怎么写", None, "hr_people"),
    ("做一套新员工安全培训课件", None, "training"),
    ("写一页纸产品介绍", None, "product"),
    ("做一份融资BP", None, "bp"),
    ("帮我写比赛的作品简介，400字以内", None, "competition"),
    ("写一份关于软件开发的技术方案", None, "bid"),
    ("对比一下三家云服务商", None, "competitor"),
    ("做个笔记本电脑采购比价", None, "procurement"),
    ("写一份差旅费用报销说明", None, "finance"),
    ("客户投诉物流太慢，帮我回复", None, "customer"),
    ("做一份北京出差三天的行程安排", None, "schedule"),
    ("写一份消防应急预案", None, "emergency"),
    ("写一份服务器宕机的故障复盘报告", None, "retro"),
    ("设计一份员工满意度调查问卷", None, "survey"),
    ("给新品写5条朋友圈文案", None, "copy"),
    ("帮我优化一下简历", None, "resume"),
    ("写软著的用户操作说明书", None, "ip"),
    ("帮我写论文摘要", "summarize", "academic"),
    ("写一篇培训心得体会，1500字", None, "reflection"),
    ("解读一下这份政策文件", None, "policy_interp"),
    ("把这些数据整理成表格", None, "table"),
    ("把这段话润色得更正式一些", "polish", None),
    ("校对一下这份通知有没有错别字", "proofread", "notice"),
    ("把这份会议纪要翻译成英文", "translate", "minutes"),
    ("把这段压缩到200字以内", "condense", None),
    ("AetherGuard 是什么？", "qa", "explainer"),
    ("这个数字能对外说吗？", "qa", None),
    ("帮我想几个活动主题", "outline", None),
    ("给我列个提纲", "outline", None),
    ("你好", None, None),
    ("谢谢", None, None),
]


class RoutingTest(unittest.TestCase):
    def test_route_table(self):
        wrong = []
        for message, operation_id, genre_id in ROUTES:
            operation, genre = prompting.route(message)
            got = (operation.id if operation else None, genre.id if genre else None)
            if got != (operation_id, genre_id):
                wrong.append(f"{message} -> {got}, expected {(operation_id, genre_id)}")
        self.assertEqual(wrong, [])

    def test_every_genre_has_a_routing_case(self):
        covered = {genre for _, _, genre in ROUTES if genre}
        self.assertEqual({preset.id for preset in GENRES} - covered, set())

    def test_follow_up_keeps_genre_from_recent_requests_or_task(self):
        operation, genre = prompting.route("把第二部分改短一点", recent_requests=["帮我把今天的会议笔记整理成纪要"])
        self.assertEqual((operation.id, genre.id), ("condense", "minutes"))
        _, genre = prompting.route("开始吧", {"title": "三月采购比价", "deliverable": "", "description": ""})
        self.assertEqual(genre.id, "procurement")
        _, genre = prompting.route("再改改", latest_title="2025 年度工作总结")
        self.assertEqual(genre.id, "summary")

    def test_current_message_outweighs_older_context(self):
        _, genre = prompting.route("改成一封邮件发给客户", recent_requests=["写一份放假通知"])
        self.assertEqual(genre.id, "email")


class PresetLibraryTest(unittest.TestCase):
    def test_presets_are_well_formed(self):
        self.assertEqual(len(BY_ID), len(ALL))
        self.assertGreaterEqual(len(GENRES), 35)
        self.assertGreaterEqual(len(OPERATIONS), 8)
        for preset in ALL:
            with self.subTest(preset=preset.id):
                self.assertIn(preset.kind, {"operation", "genre"})
                self.assertIn(preset.tier, prompting.TIER_ORDER)
                self.assertTrue(preset.keywords)
                for pattern in preset.keywords:
                    re.compile(pattern)
                self.assertTrue(1 <= len(preset.checklist) <= 6)
                self.assertTrue(all(0 < len(point) <= 120 for point in preset.checklist))
                self.assertTrue(preset.guide.startswith("适用："))
                self.assertLessEqual(len(preset.guide), 2600)

    def test_high_stakes_genres_require_critical_tier(self):
        for key in ("official", "contract_review", "contract_draft", "hr_people", "finance", "procurement", "proposal",
                    "bid", "press", "product", "competition", "emergency", "ip", "policy", "bp"):
            self.assertEqual(BY_ID[key].tier, "critical", key)

    def test_checklist_merges_operation_and_genre(self):
        rules = prompting.checklist(BY_ID["translate"], BY_ID["contract_review"])
        self.assertEqual(rules["类型"], "翻译与外文写作、合同审阅")
        self.assertEqual(rules["最低分级"], "critical")
        self.assertEqual(rules["验收要点"][:3], list(BY_ID["translate"].checklist))
        self.assertLessEqual(len(rules["验收要点"]), 12)
        self.assertIsNone(prompting.checklist(None, None))


class PromptSizeTest(unittest.TestCase):
    def test_pipeline_prompts_stay_under_desktop_limit(self):
        # desktop/agent.cjs sanitizePipelineConfig 拒绝超过 8000 字的阶段提示词，整套智能分工会失效。
        for name, value in pipeline.PROMPTS.items():
            self.assertLessEqual(len(value), 7500, name)
        self.assertIn("写作规范.最低分级", pipeline.PROMPTS["understand"])
        self.assertIn("写作规范.验收要点", pipeline.PROMPTS["accept"])

    def test_every_system_combination_stays_bounded(self):
        largest = 0
        for mode, operation, genre in itertools.product(prompting.MODE_HELP, (None, *OPERATIONS), (None, *GENRES)):
            largest = max(largest, len(prompting.conversation_system(mode, operation, genre)))
        for kind, operation, genre in itertools.product(("docx", "pptx"), (None, *OPERATIONS), (None, *GENRES)):
            largest = max(largest, len(prompting.plan_system(kind, operation, genre)))
        self.assertLessEqual(largest, prompting.MAX_SYSTEM)

    def test_system_keeps_output_protocol_and_injects_only_matched_presets(self):
        system = prompting.conversation_system("pptx", None, BY_ID["minutes"])
        self.assertIn('{"action":"reply","reply":"给员工的 Markdown 答复"', system)
        self.assertIn('"kind":"docx"', system)
        self.assertIn('"kind":"pptx"', system)
        self.assertIn("## 交付形式判定", system)
        self.assertIn("本轮目标是制作或修改 PPT", system)
        self.assertIn("## PPT 咨询/发布会级自由画布设计规范", system)
        self.assertIn("## 回复排版", system)
        self.assertNotIn("## Word 导出规则", system)
        self.assertIn("## 写作规范：会议纪要", system)
        self.assertEqual(system.count("## 写作规范："), 1)
        bare = prompting.conversation_system("reply")
        self.assertNotIn("## 写作规范：", bare)
        self.assertIn("本轮只作文字回复，action 必须是 reply。", bare)
        self.assertNotIn("自由画布设计规范", bare)
        plan = prompting.plan_system("docx", BY_ID["polish"], BY_ID["notice"])
        self.assertIn('{"title":"标题"', plan)
        self.assertNotIn('"action"', plan)
        self.assertEqual(plan.count("## 写作规范："), 2)


if __name__ == "__main__":
    unittest.main()
