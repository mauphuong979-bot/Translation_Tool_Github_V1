"""Tests for ai_translate_api with a fake client (no network, no cost).
Run from the scratch folder: python -m unittest test_ai_translate_api.py"""
import json
import threading
import unittest
from types import SimpleNamespace

import test_review_roundtrip as fixtures
from ai_translate_api import (FALLBACK_BETA, AIConfig, build_payload, build_system_prompt, describe_api_error, estimate_cost,
                              find_company_names, load_config, make_batches, title_case_name, translate_items, usage_cost)
from review_roundtrip import apply_ai_results, preview_ai_results, read_review_manifest

saved = fixtures.saved


def message(items=None, stop_reason="end_turn", text=None, usage=(100, 50)):
    body = text if text is not None else json.dumps({"items": items or []}, ensure_ascii=False)
    return SimpleNamespace(stop_reason=stop_reason, content=[SimpleNamespace(type="text", text=body)],
                           usage=SimpleNamespace(input_tokens=usage[0], output_tokens=usage[1],
                                                 cache_read_input_tokens=0, cache_creation_input_tokens=0))


def echo(payload, translate=lambda entry: "EN: " + entry["source"]):
    return message([{"review_id": entry["review_id"], "translation": translate(entry), "highlight": [], "note": ""}
                    for entry in payload])


class FakeStream:
    def __init__(self, result):
        self.result = result

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def get_final_message(self):
        return self.result


class FakeMessages:
    def __init__(self, client, kind):
        self.client, self.kind = client, kind

    def stream(self, **request):
        with self.client.lock:
            self.client.requests.append((self.kind, request))
        if "tools" in request:
            return FakeStream(self.client.lookup(request))
        return FakeStream(self.client.respond(json.loads(request["messages"][0]["content"])))

    def count_tokens(self, **request):
        self.client.counted.append(request)
        return SimpleNamespace(input_tokens=1000)


def lookup_reply(found, searches=1, stop_reason="end_turn"):
    """A lookup response: {name: (international, url)} as the final JSON text, with web search usage."""
    body = json.dumps({"names": [{"vietnamese": name, "international": international, "source_url": url}
                                 for name, (international, url) in found.items()]}, ensure_ascii=False)
    reply = message(text="Kết quả: " + body, stop_reason=stop_reason)
    reply.usage.server_tool_use = SimpleNamespace(web_search_requests=searches)
    return reply


def no_lookup(request):
    raise AssertionError("unexpected lookup request")


class FakeClient:
    def __init__(self, respond=echo, lookup=no_lookup):
        self.respond, self.lookup, self.requests, self.counted, self.lock = respond, lookup, [], [], threading.Lock()
        self.messages = FakeMessages(self, "messages")
        self.beta = SimpleNamespace(messages=FakeMessages(self, "beta"))


def item(index, original="Doanh thu thuần", references=None):
    return {"review_id": f"B-C{index:03d}", "original": original,
            "references": references if references is not None else
            [{"source": "dictionary", "score": 0.9, "original": "Doanh thu", "translation": "Revenue"}]}


CONFIG = AIConfig(api_key="test", max_workers=2)


class PayloadAndPromptTests(unittest.TestCase):
    def test_payload_is_allowlisted(self):
        leaky = dict(item(1), report_text="SECRET PARAGRAPH", comment_text="SECRET COMMENT", comment_author="Reviewer",
                     anchor_name="TT_AI_X", formatting={"size": 10})
        leaky["references"][0]["extra"] = "SECRET REF"
        payload = build_payload([leaky])
        self.assertEqual(set(payload[0]), {"review_id", "source", "references"})
        self.assertEqual(set(payload[0]["references"][0]), {"label", "vietnamese", "translation"})
        self.assertEqual(payload[0]["references"][0]["label"], "Từ điển (90%)")
        self.assertNotIn("SECRET", json.dumps(payload, ensure_ascii=False))

    def test_report_to_request_sends_only_ai_form_content(self):
        report, worksheet, items = fixtures.ReviewRoundtripTests().pair(count=2)
        report.add_paragraph("Đoạn khác của báo cáo, không được gửi đi")
        manifest = read_review_manifest(saved(report))
        client = FakeClient()
        translate_items(manifest["items"], manifest["target_col"], CONFIG, client)
        sent = "".join(request["messages"][0]["content"] for _, request in client.requests)
        for forbidden in ("Đoạn khác", "Reference suggestion", "Translation Tool", "TT_AI_"):
            self.assertNotIn(forbidden, sent)
        sent_ids = [entry["review_id"] for _, request in client.requests for entry in json.loads(request["messages"][0]["content"])]
        ai_form_ids = [table.cell(1, 0).text.strip() for table in worksheet.tables]
        self.assertEqual(sorted(sent_ids), sorted(ai_form_ids))

    def test_system_prompt_for_each_language(self):
        expectations = {"E": "Dec 31, 2025", "Hs": "giản thể", "Ht": "phồn thể"}
        for target, marker in expectations.items():
            with self.subTest(target=target):
                prompt = build_system_prompt(target)
                self.assertIn(marker, prompt)
                for shared in ("Dịch đầy đủ", "Tên riêng – tên người", "Thẻ mẫu"):
                    self.assertIn(shared, prompt)
                self.assertNotIn("file Word", prompt)
        with self.assertRaises(ValueError):
            build_system_prompt("V")

    def test_batches_respect_count_and_size(self):
        items = [item(index) for index in range(7)]
        self.assertEqual([len(batch) for batch in make_batches(items, 3, 100000)], [3, 3, 1])
        long_items = [item(index, original="x" * 400) for index in range(4)]
        item_chars = len(json.dumps(build_payload([long_items[0]]), ensure_ascii=False))
        self.assertEqual([len(batch) for batch in make_batches(long_items, 25, item_chars * 2 + 10)], [2, 2])
        self.assertEqual([len(batch) for batch in make_batches([item(1, original="x" * 5000)], 25, 1000)], [1])

    def test_load_config(self):
        self.assertIsNone(load_config(None))
        self.assertIsNone(load_config({"model": "claude-opus-5-5"}))
        config = load_config({"api_key": "k", "model": "claude-sonnet-5-5", "batch_size": 10, "unknown": 1})
        self.assertEqual((config.model, config.batch_size, config.effort), ("claude-sonnet-5-5", 10, "medium"))


class TranslateTests(unittest.TestCase):
    def test_request_shape_with_and_without_fallbacks(self):
        client = FakeClient()
        translate_items([item(1)], "E", CONFIG, client)
        kind, request = client.requests[0]
        self.assertEqual(kind, "beta")
        self.assertEqual((request["model"], request["betas"], request["fallbacks"]), ("claude-opus-5-5", [FALLBACK_BETA], "default"))
        self.assertEqual(request["output_config"]["effort"], "medium")
        self.assertEqual(request["output_config"]["format"]["type"], "json_schema")
        self.assertEqual(request["system"][0]["cache_control"], {"type": "ephemeral"})
        client = FakeClient()
        translate_items([item(1)], "E", AIConfig(api_key="test", fallbacks=False), client)
        kind, request = client.requests[0]
        self.assertEqual(kind, "messages")
        self.assertNotIn("fallbacks", request)
        self.assertNotIn("betas", request)

    def test_results_in_order_with_notes_progress_and_usage(self):
        items = [item(index, original=f"Câu {index}") for index in range(5)]
        progress = []

        def respond(payload):
            return message([{"review_id": entry["review_id"], "translation": "T " + entry["source"],
                             "highlight": ["T"], "note": "Cần xác nhận" if entry["review_id"] == "B-C002" else ""}
                            for entry in reversed(payload)])

        outcome = translate_items(items, "E", AIConfig(api_key="test", batch_size=2), FakeClient(respond),
                                  progress=lambda done, total: progress.append((done, total)))
        self.assertEqual([result["review_id"] for result in outcome["results"]], [entry["review_id"] for entry in items])
        self.assertEqual(outcome["results"][0], {"review_id": "B-C000", "translation": "T Câu 0", "highlight": ["T"]})
        self.assertEqual(outcome["notes"], {"B-C002": "Cần xác nhận"})
        self.assertEqual(progress[-1], (5, 5))
        self.assertEqual(outcome["usage"]["input_tokens"], 300)
        self.assertEqual((outcome["errors"], outcome["failed_ids"], outcome["warnings"]), ([], [], []))

    def test_max_tokens_splits_batch(self):
        def respond(payload):
            return message(stop_reason="max_tokens", text="{") if len(payload) > 1 else echo(payload)

        client = FakeClient(respond)
        outcome = translate_items([item(index) for index in range(4)], "E", AIConfig(api_key="test", max_workers=1), client)
        self.assertEqual(len(outcome["results"]), 4)
        self.assertEqual(len(client.requests), 7)  # 4 -> 2+2 -> 1+1+1+1
        self.assertEqual(outcome["usage"]["input_tokens"], 700)

    def test_failed_batch_is_reported_and_others_kept(self):
        def respond(payload):
            if payload[0]["review_id"] == "B-C000":
                return message(stop_reason="refusal", text="")
            return echo(payload)

        outcome = translate_items([item(index) for index in range(4)], "E", AIConfig(api_key="test", batch_size=2), FakeClient(respond))
        self.assertEqual(outcome["failed_ids"], ["B-C000", "B-C001"])
        self.assertEqual([result["review_id"] for result in outcome["results"]], ["B-C002", "B-C003"])
        self.assertIn("từ chối", outcome["errors"][0])

    def test_invalid_json_unknown_and_missing_ids(self):
        outcome = translate_items([item(1)], "E", CONFIG, FakeClient(lambda payload: message(text="not json")))
        self.assertEqual(outcome["failed_ids"], ["B-C001"])

        def respond(payload):
            return message([{"review_id": "OTHER-C999", "translation": "x", "highlight": [], "note": ""},
                            {"review_id": payload[0]["review_id"], "translation": "ok", "highlight": [], "note": ""}])

        outcome = translate_items([item(1), item(2)], "E", CONFIG, FakeClient(respond))
        self.assertEqual([result["review_id"] for result in outcome["results"]], ["B-C001"])
        self.assertTrue(any("OTHER-C999" in warning for warning in outcome["warnings"]))
        self.assertTrue(any("B-C002" in warning for warning in outcome["warnings"]))

    def test_end_to_end_with_review_roundtrip(self):
        report, _, items = fixtures.ReviewRoundtripTests().pair(count=3, target="E")
        report_bytes = saved(report)
        manifest = read_review_manifest(report_bytes)

        def respond(payload):
            if payload[0]["review_id"] == items[0]["review_id"]:
                return message(stop_reason="refusal", text="")
            return echo(payload, lambda entry: "Charter capital VND 1,234.56.")

        outcome = translate_items(manifest["items"], "E", AIConfig(api_key="test", batch_size=1), FakeClient(respond))
        preview = preview_ai_results(report_bytes, outcome["results"])
        self.assertEqual(preview["counts"], {"ready": 2, "missing": 1})
        output, summary = apply_ai_results(report_bytes, outcome["results"])
        self.assertEqual(summary["applied"], 2)
        remaining = read_review_manifest(output.getvalue())["items"]
        self.assertEqual([entry["review_id"] for entry in remaining], [items[0]["review_id"]])


ABC = "Công ty TNHH Thương mại và Dịch vụ ABC"
WEB = AIConfig(api_key="test", web_search=True, max_workers=1)


class CompanyLookupTests(unittest.TestCase):
    def test_find_company_names(self):
        cases = {
            ABC + " có trụ sở tại Hà Nội.": [ABC],
            "Công ty có vốn điều lệ 1.234,56 đồng.": [],
            "Khách hàng là CÔNG TY CỔ PHẦN XYZ VIỆT NAM, theo hợp đồng.": ["CÔNG TY CỔ PHẦN XYZ VIỆT NAM"],
            "Vay Ngân hàng TMCP Ngoại thương Việt Nam (Vietcombank) và Công ty Mẹ.": ["Ngân hàng TMCP Ngoại thương Việt Nam"],
            "Công ty Cổ phần Đầu tư Xây dựng Hòa Bình và các công ty con": ["Công ty Cổ phần Đầu tư Xây dựng Hòa Bình"],
            "Công ty mẹ là Công ty TNHH ABC Holdings.": ["Công ty TNHH ABC Holdings"],
            ABC + " và Công ty Cổ phần XYZ Việt Nam ký hợp đồng.": [ABC, "Công ty Cổ phần XYZ Việt Nam"],
            "Công ty TNHH ABC ký hợp đồng": ["Công ty TNHH ABC"],
            "Công ty Cổ phần Bảo hiểm có trụ sở": ["Công ty Cổ phần Bảo hiểm"],
            "Công ty TNHH Thương mại.": ["Công ty TNHH Thương mại"],
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(find_company_names([{"original": text}]), expected)
        self.assertEqual(find_company_names([{"original": ABC + "."}, {"original": ABC + ", chi nhánh"}]), [ABC])

    def test_lookup_sends_only_names_and_translation_uses_found_names(self):
        items = [item(1, original=ABC), item(2, original="Doanh thu thuần")]
        client = FakeClient(lookup=lambda request: lookup_reply({ABC: ("ABC Trading and Services Co., Ltd", "https://masothue.com/0101234567-abc")}, 2))
        outcome = translate_items(items, "E", WEB, client)
        lookups = [request for _, request in client.requests if "tools" in request]
        translations = [request for _, request in client.requests if "tools" not in request]
        self.assertEqual(len(lookups), 1)
        self.assertEqual(json.loads(lookups[0]["messages"][0]["content"]), [ABC])
        tool = lookups[0]["tools"][0]
        self.assertEqual((tool["type"], tool["allowed_domains"], tool["max_uses"]),
                         ("web_search_20260209", ["masothue.com", "thuvienphapluat.vn"], 6))
        self.assertNotIn("format", lookups[0]["output_config"])
        self.assertNotIn("Doanh thu", lookups[0]["messages"][0]["content"])
        payload = json.loads(translations[0]["messages"][0]["content"])
        self.assertEqual(payload[0]["company_names"], [{"vietnamese": ABC, "international": "ABC Trading and Services Co., Ltd"}])
        self.assertNotIn("company_names", payload[1])
        self.assertEqual(outcome["company_names"][ABC]["international"], "ABC Trading and Services Co., Ltd")
        self.assertEqual(outcome["usage"]["web_search_requests"], 2)

    def test_untrusted_lookup_answers_are_dropped(self):
        found = {ABC: ("Wrong Source Ltd", "https://example.com/abc"),
                 "Công ty TNHH Khác": ("Other Ltd", "https://masothue.com/x"),
                 "Công ty TNHH Trống": ("", "https://masothue.com/y")}
        outcome = translate_items([item(1, original=ABC + ".")], "E", WEB, FakeClient(lookup=lambda request: lookup_reply(found)))
        self.assertEqual(outcome["company_names"], {})
        sub_domain = translate_items([item(1, original=ABC + ".")], "E", WEB,
                                     FakeClient(lookup=lambda request: lookup_reply({ABC: ("ABC Ltd", "https://www.thuvienphapluat.vn/abc")})))
        self.assertIn(ABC, sub_domain["company_names"])

    def test_title_case_name(self):
        cases = {
            "ABC TRADING AND SERVICES COMPANY LIMITED": "Abc Trading and Services Company Limited",
            "HOA BINH CONSTRUCTION GROUP JSC": "Hoa Binh Construction Group JSC",
            "NEW INCREASE (VN) CO., LTD": "New Increase (VN) Co., Ltd",
            "THE BANK FOR FOREIGN TRADE OF VIETNAM": "The Bank for Foreign Trade of Vietnam",
            "TRANS-PACIFIC LOGISTICS II LLC": "Trans-Pacific Logistics II LLC",
            "ABC Vietnam Co., Ltd": "ABC Vietnam Co., Ltd",
            "": "",
        }
        for name, expected in cases.items():
            with self.subTest(name=name):
                self.assertEqual(title_case_name(name), expected)

    def test_all_caps_lookup_answer_is_title_cased(self):
        outcome = translate_items([item(1, original=ABC + ".")], "E", WEB, FakeClient(
            lookup=lambda request: lookup_reply({ABC: (" ABC  TRADING CO., LTD ", "https://masothue.com/abc")})))
        self.assertEqual(outcome["company_names"][ABC]["international"], "Abc Trading Co., Ltd")

    def test_lookup_answer_matches_name_written_differently_in_word(self):
        source = "Công ty TNHH Chỉ May American & Efird​ Việt Nam"
        echoed = "CÔNG TY TNHH CHỈ MAY AMERICAN & EFIRD VIỆT NAM"
        outcome = translate_items([item(1, original=source)], "E", WEB, FakeClient(
            lookup=lambda request: lookup_reply({echoed: ("AMERICAN & EFIRD VIETNAM THREAD LIMITED LIABILITY COMPANY",
                                                          "https://masothue.com/3603520029")})))
        self.assertEqual([entry["international"] for entry in outcome["company_names"].values()],
                         ["American & Efird Vietnam Thread Limited Liability Company"])

    def test_pause_turn_is_continued(self):
        replies = [lookup_reply({}, stop_reason="pause_turn"), lookup_reply({ABC: ("ABC Ltd", "https://masothue.com/abc")})]
        client = FakeClient(lookup=lambda request: replies.pop(0))
        outcome = translate_items([item(1, original=ABC + ".")], "E", WEB, client)
        lookups = [request for _, request in client.requests if "tools" in request]
        self.assertEqual(len(lookups), 2)
        self.assertEqual([entry["role"] for entry in lookups[1]["messages"]], ["user", "assistant"])
        self.assertIn(ABC, outcome["company_names"])
        self.assertEqual(outcome["usage"]["web_search_requests"], 2)

    def test_failed_lookup_does_not_stop_translation(self):
        outcome = translate_items([item(1, original=ABC + ".")], "E", WEB,
                                  FakeClient(lookup=lambda request: message(text="không có JSON")))
        self.assertEqual(len(outcome["results"]), 1)
        self.assertEqual(outcome["company_names"], {})
        self.assertTrue(any("Tra cứu tên công ty" in warning for warning in outcome["warnings"]))

    def test_no_lookup_when_disabled_or_no_names(self):
        translate_items([item(1, original=ABC + ".")], "E", CONFIG, FakeClient())
        translate_items([item(1)], "E", WEB, FakeClient())

    def test_saved_names_are_used_and_not_looked_up(self):
        other = "Công ty Cổ phần XYZ Việt Nam"
        items = [item(1, original=ABC), item(2, original=other + "."), item(3, original="Doanh thu thuần")]
        client = FakeClient(lookup=lambda request: lookup_reply({other: ("XYZ Vietnam JSC", "https://masothue.com/xyz")}))
        outcome = translate_items(items, "E", WEB, client, saved_names={ABC: "ABC Trading Ltd", "Công ty TNHH Không Liên Quan": "Unrelated Ltd"})
        lookups = [request for _, request in client.requests if "tools" in request]
        self.assertEqual(json.loads(lookups[0]["messages"][0]["content"]), [other])
        self.assertEqual(outcome["searched_names"], [other])
        self.assertEqual(outcome["company_names"][ABC], {"international": "ABC Trading Ltd", "source_url": "", "saved": True})
        self.assertNotIn("Công ty TNHH Không Liên Quan", outcome["company_names"])
        payload = [entry for _, request in client.requests if "tools" not in request
                   for entry in json.loads(request["messages"][0]["content"])]
        self.assertEqual(payload[0]["company_names"], [{"vietnamese": ABC, "international": "ABC Trading Ltd"}])
        self.assertEqual(payload[1]["company_names"], [{"vietnamese": other, "international": "XYZ Vietnam JSC"}])
        self.assertNotIn("company_names", payload[2])

    def test_names_inside_paragraph_are_not_looked_up(self):
        other = "Công ty Cổ phần XYZ Việt Nam"
        sentence = item(1, original=ABC + " và " + other + " ký hợp đồng.")
        client = FakeClient()  # no_lookup raises if a lookup is sent
        outcome = translate_items([sentence], "E", WEB, client, saved_names={ABC: "ABC Trading Ltd"})
        self.assertEqual((outcome["searched_names"], outcome["company_names"]), ([], {}))
        self.assertNotIn("company_names", json.loads(client.requests[0][1]["messages"][0]["content"])[0])
        self.assertEqual(estimate_cost([sentence], "E", WEB, FakeClient())["company_names"], [])
        upper = item(2, original="  CÔNG TY CỔ PHẦN XYZ VIỆT NAM  ")
        self.assertEqual(estimate_cost([upper], "E", WEB, FakeClient())["company_names"], ["CÔNG TY CỔ PHẦN XYZ VIỆT NAM"])

    def test_saved_names_work_without_web_search(self):
        client = FakeClient()
        outcome = translate_items([item(1, original=ABC + ".")], "E", CONFIG, client, saved_names={ABC: "ABC Ltd"})
        self.assertEqual(outcome["searched_names"], [])
        payload = json.loads(client.requests[0][1]["messages"][0]["content"])
        self.assertEqual(payload[0]["company_names"], [{"vietnamese": ABC, "international": "ABC Ltd"}])
        estimate = estimate_cost([item(1, original=ABC + ".")], "E", WEB, FakeClient(), saved_names={ABC: "ABC Ltd"})
        self.assertEqual((estimate["company_names"], estimate["web_searches_estimate"]), ([], 0))

    def test_estimate_and_cost_include_searches(self):
        estimate = estimate_cost([item(1, original=ABC + ".")], "E", WEB, FakeClient())
        self.assertEqual((estimate["company_names"], estimate["web_searches_estimate"]), ([ABC], 1))
        self.assertAlmostEqual(usage_cost({"web_search_requests": 3}, "claude-opus-5-5"), 0.03)


class CostTests(unittest.TestCase):
    def test_estimate_uses_count_tokens(self):
        client = FakeClient()
        estimate = estimate_cost([item(index) for index in range(3)], "Ht", AIConfig(api_key="test", batch_size=2), client)
        self.assertEqual(len(client.counted), 2)
        self.assertEqual(estimate["input_tokens"], 2000)
        self.assertGreater(estimate["usd_estimate"], 0)
        self.assertNotIn("fallbacks", client.counted[0])

    def test_usage_cost(self):
        cost = usage_cost({"input_tokens": 1_000_000, "output_tokens": 1_000_000, "cache_read_input_tokens": 1_000_000}, "claude-opus-5-5")
        self.assertAlmostEqual(cost, 24.20)
        self.assertIsNone(usage_cost({}, "unknown-model"))



def api_error(status, text="boom"):
    import anthropic
    import httpx

    response = httpx.Response(status, request=httpx.Request("POST", "https://api.anthropic.com/v1/messages"))
    classes = {400: anthropic.BadRequestError, 401: anthropic.AuthenticationError, 429: anthropic.RateLimitError}
    error_class = classes.get(status, anthropic.InternalServerError if status >= 500 else anthropic.APIStatusError)
    return error_class(text, response=response, body=None)


class ErrorMessageTests(unittest.TestCase):
    def test_api_errors_are_explained_in_vietnamese(self):
        import anthropic
        import httpx

        self.assertIn("Khóa API không hợp lệ", describe_api_error(api_error(401, "Invalid bearer token")))
        self.assertIn("(mã lỗi 401)", describe_api_error(api_error(401)))
        self.assertIn("Vượt giới hạn", describe_api_error(api_error(429)))
        self.assertIn("quá tải", describe_api_error(api_error(529)))
        self.assertIn("sự cố", describe_api_error(api_error(500)))
        self.assertIn("hết tín dụng", describe_api_error(api_error(400, "Your credit balance is too low")))
        connection = anthropic.APIConnectionError(request=httpx.Request("POST", "https://api.anthropic.com"))
        self.assertIn("Không kết nối được", describe_api_error(connection))
        self.assertEqual(describe_api_error(ValueError("AI từ chối dịch lô này")), "AI từ chối dịch lô này")

    def test_same_error_in_every_batch_is_reported_once(self):
        def respond(payload):
            raise api_error(401, "Invalid bearer token")

        outcome = translate_items([item(index) for index in range(4)], "E", AIConfig(api_key="test", batch_size=1), FakeClient(respond))
        self.assertEqual(len(outcome["errors"]), 1)
        self.assertTrue(outcome["errors"][0].startswith("Không dịch được 4 mục (từ B-C000): Khóa API không hợp lệ"))
        self.assertNotIn("bearer", outcome["errors"][0])
        self.assertEqual(len(outcome["error_details"]), 4)
        self.assertIn("Invalid bearer token", outcome["error_details"][0])


if __name__ == "__main__":
    unittest.main()
