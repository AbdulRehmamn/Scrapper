import io
import unittest
from unittest.mock import AsyncMock, patch
from urllib.parse import urlsplit

from openpyxl import load_workbook
from pydantic import ValidationError

import app


SAMPLE_HTML = """
<!doctype html>
<html>
  <head>
    <title>Contractor Directory</title>
    <meta name="description" content="Approved contractor directory">
    <script type="application/ld+json">
      {"@context":"https://schema.org","@type":"Organization","name":"Schema Builders",
       "email":"hello@schema.example","telephone":"+1 415 555 0100",
       "url":"https://schema.example","address":{"streetAddress":"10 Market Street",
       "addressLocality":"San Francisco","addressRegion":"CA","postalCode":"94105"}}
    </script>
  </head>
  <body>
    <h1>Approved Contractors</h1>
    <article class="listing-card contractor">
      <h2>Alpha Electric</h2><p>Call +1 (212) 555-0100 or alpha@example.com</p>
      <a href="https://alpha.example">Website</a>
    </article>
    <article class="listing-card contractor">
      <h2>Beta Mechanical</h2><p>Call +1 (312) 555-0110 or beta@example.com</p>
      <a href="/contractors/beta">Profile</a>
    </article>
    <a href="/directory?page=2&utm_source=test">Next page</a>
    <a href="https://linkedin.com/company/example">LinkedIn</a>
    <table><caption>Licenses</caption><thead><tr><th>Company</th><th>License</th></tr></thead>
      <tbody><tr><td>Alpha Electric</td><td>C-10</td></tr><tr><td>Beta Mechanical</td><td>C-20</td></tr></tbody>
    </table>
  </body>
</html>
"""


class UrlSafetyTests(unittest.TestCase):
    def test_adds_https_to_plain_hostname(self):
        self.assertEqual(
            app.normalize_input_url("example.com/directory"),
            "https://example.com/directory",
        )

    def test_blocks_local_and_private_destinations(self):
        for value in (
            "http://localhost/",
            "http://127.0.0.1/",
            "http://10.0.0.1/",
            "http://192.168.1.2/",
            "http://[::1]/",
            "http://internal.local/",
        ):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    app.normalize_input_url(value)

    def test_strips_tracking_query_and_fragments(self):
        value = app.canonicalize_link(
            "https://example.com/start",
            "/directory?page=2&utm_source=test#results",
        )
        self.assertEqual(value, "https://example.com/directory?page=2")

    def test_accepts_five_thousand_pages_but_not_more(self):
        config = app.ScrapeRequest(url="https://example.com", max_pages=5_000)
        self.assertEqual(config.max_pages, 5_000)
        with self.assertRaises(ValidationError):
            app.ScrapeRequest(url="https://example.com", max_pages=5_001)


class ExtractionTests(unittest.TestCase):
    def config(self, **changes):
        values = {
            "url": "https://example.com/directory",
            "mode": "smart",
            "max_pages": 1,
            "max_depth": 0,
            "delay_ms": 200,
            "render_javascript": False,
            "respect_robots": True,
            "selector": "",
            "include_text": "",
            "exclude_text": "",
        }
        values.update(changes)
        return app.ScrapeRequest(**values)

    def test_smart_mode_extracts_records_contacts_tables_and_links(self):
        result = app.extract_page(SAMPLE_HTML, "https://example.com/directory", self.config())
        titles = {record["title"] for record in result["records"]}
        contact_values = {item["normalized"] for item in result["contacts"]}
        self.assertIn("Schema Builders", titles)
        self.assertIn("Alpha Electric", titles)
        self.assertIn("Beta Mechanical", titles)
        self.assertIn("alpha@example.com", contact_values)
        self.assertIn("+12125550100", contact_values)
        self.assertEqual(result["tables"][0]["title"], "Licenses")
        self.assertEqual(len(result["tables"][0]["rows"]), 2)
        self.assertIn("https://example.com/directory?page=2", result["links"])

    def test_custom_selector_creates_one_record_per_match(self):
        config = self.config(mode="custom", selector="article.listing-card")
        result = app.extract_page(SAMPLE_HTML, "https://example.com/directory", config)
        self.assertEqual(len(result["records"]), 2)
        self.assertEqual(result["records"][0]["title"], "Alpha Electric")

    def test_contacts_mode_omits_tables_and_records(self):
        config = self.config(mode="contacts")
        result = app.extract_page(SAMPLE_HTML, "https://example.com/directory", config)
        self.assertEqual(result["records"], [])
        self.assertEqual(result["tables"], [])
        self.assertGreaterEqual(len(result["contacts"]), 4)

    def test_email_phone_mode_keeps_paired_contact_records(self):
        config = self.config(mode="contact_numbers")
        result = app.extract_page(SAMPLE_HTML, "https://example.com/directory", config)
        titles = {record["title"] for record in result["records"]}
        self.assertIn("Schema Builders", titles)
        self.assertIn("Alpha Electric", titles)
        self.assertIn("Beta Mechanical", titles)
        self.assertTrue(result["contacts"])
        self.assertLessEqual(
            {contact["type"] for contact in result["contacts"]},
            {"Email", "Phone"},
        )


class ExportTests(unittest.TestCase):
    def result(self):
        extracted = app.extract_page(
            SAMPLE_HTML,
            "https://example.com/directory",
            app.ScrapeRequest(url="https://example.com", max_pages=1, max_depth=0),
        )
        return {
            "summary": {"start_url": "https://example.com", "records": len(extracted["records"])},
            "records": [{"id": index + 1, **record} for index, record in enumerate(extracted["records"])],
            "contacts": extracted["contacts"],
            "tables": extracted["tables"],
            "pages": [extracted["page"]],
            "errors": [],
        }

    def test_csv_export_is_utf8_and_contains_records(self):
        payload = app.csv_bytes(self.result()).decode("utf-8-sig")
        self.assertIn("Schema Builders", payload)
        self.assertIn("Source URL", payload)

    def test_xlsx_export_has_expected_sheets(self):
        workbook = load_workbook(io.BytesIO(app.xlsx_bytes(self.result())), read_only=True)
        self.assertEqual(
            workbook.sheetnames,
            ["Summary", "Records", "Contacts", "Pages", "Tables", "Errors"],
        )
        self.assertGreater(sum(1 for _ in workbook["Records"].iter_rows()), 1)


class JobPipelineTests(unittest.IsolatedAsyncioTestCase):
    async def test_scrape_job_completes_and_builds_summary(self):
        config = app.ScrapeRequest(
            url="https://example.com/directory",
            max_pages=1,
            max_depth=0,
            delay_ms=200,
            respect_robots=False,
        )
        job = app.Job(id="test-job", config=config)
        fetched = app.FetchResult(
            url=config.url,
            html=SAMPLE_HTML,
            status_code=200,
            content_type="text/html",
        )
        with (
            patch("app.validate_public_url", new=AsyncMock(side_effect=lambda value: value)),
            patch("app.fetch_static", new=AsyncMock(return_value=fetched)),
        ):
            await app.run_scrape_job(job)

        self.assertEqual(job.status, "completed")
        self.assertEqual(job.progress, 100)
        self.assertEqual(job.result["summary"]["pages_scraped"], 1)
        self.assertGreaterEqual(job.result["summary"]["records"], 3)
        self.assertGreaterEqual(job.result["summary"]["emails"], 3)

        public = job.public()
        self.assertIn("summary", public)
        self.assertNotIn("result", public)

    async def test_mocked_job_processes_five_thousand_pages(self):
        async def fetch_page(_client, url):
            page_number = int(urlsplit(url).path.rsplit("/", 1)[-1])
            next_link = (
                f'<a href="/pages/{page_number + 1}">Next</a>'
                if page_number < 5_000 else ""
            )
            return app.FetchResult(
                url=url,
                html=f"<html><head><title>Page {page_number}</title></head><body>{next_link}</body></html>",
                status_code=200,
                content_type="text/html",
            )

        config = app.ScrapeRequest(
            url="https://example.com/pages/1",
            max_pages=5_000,
            max_depth=5_000,
            delay_ms=200,
            respect_robots=False,
        )
        job = app.Job(id="five-thousand-page-test", config=config)
        with (
            patch("app.validate_public_url", new=AsyncMock(side_effect=lambda value: value)),
            patch("app.fetch_static", new=AsyncMock(side_effect=fetch_page)),
            patch("app.asyncio.sleep", new=AsyncMock()),
        ):
            await app.run_scrape_job(job)

        self.assertEqual(job.status, "completed")
        self.assertEqual(job.result["summary"]["pages_scraped"], 5_000)
        self.assertEqual(len(job.result["pages"]), 5_000)


class ResultPaginationTests(unittest.TestCase):
    def test_server_side_data_pagination_and_search(self):
        config = app.ScrapeRequest(url="https://example.com", max_pages=5_000)
        job = app.Job(id="pagination-test", config=config, status="completed")
        job.result = {
            "summary": {"records": 3},
            "records": [
                {"id": 1, "title": "Alpha Electric"},
                {"id": 2, "title": "Beta Mechanical"},
                {"id": 3, "title": "Gamma Plumbing"},
            ],
            "contacts": [], "tables": [], "pages": [], "errors": [],
        }
        app.JOBS[job.id] = job
        try:
            page = app.get_job_data(job.id, "records", offset=1, limit=1, search="")
            self.assertEqual(page["count"], 3)
            self.assertEqual(page["items"][0]["title"], "Beta Mechanical")
            filtered = app.get_job_data(job.id, "records", offset=0, limit=50, search="gamma")
            self.assertEqual(filtered["count"], 1)
            self.assertEqual(filtered["items"][0]["id"], 3)
        finally:
            app.JOBS.pop(job.id, None)


class NewFeaturesTests(unittest.TestCase):
    def test_phone_state_and_timezone_resolution(self):
        state_ca, tz_pst = app.resolve_phone_state_tz("+14155550100")
        self.assertEqual(state_ca, "CA")
        self.assertEqual(tz_pst, "PST")

        state_ny, tz_est = app.resolve_phone_state_tz("+12125550100")
        self.assertEqual(state_ny, "NY")
        self.assertEqual(tz_est, "EST")

        state_tx, tz_cst = app.resolve_phone_state_tz("+17135550100")
        self.assertEqual(state_tx, "TX")
        self.assertEqual(tz_cst, "CST")

    def test_contractor_trade_classification(self):
        self.assertEqual(app.classify_contractor_trade("Cool Air HVAC Repair LLC"), "HVAC Contractor")
        self.assertEqual(app.classify_contractor_trade("Alpha Mechanical & Electrical Systems"), "MEP Contractor")
        self.assertEqual(app.classify_contractor_trade("Apex General Contractor Inc"), "GC Contractor")
        self.assertEqual(app.classify_contractor_trade("Voltage Electrical Services"), "Electrical Contractor")
        self.assertEqual(app.classify_contractor_trade("Flow Master Plumbing"), "Plumbing Contractor")

    def test_pdf_link_canonicalization_and_html_generation(self):
        pdf_link = app.canonicalize_link("https://www.dcaa.mil", "/Portals/88/FY20AllActiveContractorListing_1.pdf")
        self.assertEqual(pdf_link, "https://www.dcaa.mil/Portals/88/FY20AllActiveContractorListing_1.pdf")

        html_out = app.parse_pdf_to_html(b"%PDF-1.4 sample content", "https://example.com/test.pdf")
        self.assertIn("PDF Document: test.pdf", html_out)

    def test_phone_state_tz_and_whatsapp(self):
        phones = app.extract_phones("Call office +1 (415) 555-0100")
        states = {item["state"] for item in phones}
        self.assertIn("CA", states)

        html_wa = '<html><body><a href="https://wa.me/14155550100">WhatsApp Us</a></body></html>'
        extracted = app.extract_page(html_wa, "https://example.com", app.ScrapeRequest(url="https://example.com"))
        wa_contacts = [c for c in extracted["contacts"] if "Phone" in c["type"]]
        self.assertTrue(wa_contacts)
        self.assertEqual(wa_contacts[0]["state"], "CA")

    def test_external_domain_crawling_flag(self):
        config_same = app.ScrapeRequest(url="https://example.com", allow_external_domains=False)
        config_ext = app.ScrapeRequest(url="https://example.com", allow_external_domains=True)
        self.assertFalse(config_same.allow_external_domains)
        self.assertTrue(config_ext.allow_external_domains)

    def test_person_name_extraction(self):
        # 1. From table row dict
        row = {
            "Contractor's Name": "A Good Sign & Graphics Company",
            "Contact Person": "Richard Abedi",
            "E-Mail Address": "richard@agoodsign.com"
        }
        self.assertEqual(app.extract_person_name(row_dict=row), "Richard Abedi")

        # 2. From text label
        self.assertEqual(app.extract_person_name(text="Contact Person: Brian Connolly"), "Brian Connolly")

    def test_auto_page_detection_and_paid_website_auth(self):
        cookies = app.parse_cookies_dict("sessionid=abc123xyz; token=eyJ123")
        self.assertEqual(cookies.get("sessionid"), "abc123xyz")
        self.assertEqual(cookies.get("token"), "eyJ123")

        headers = app.parse_headers_dict("Authorization: Bearer mytoken123\nX-API-Key: test_key")
        self.assertEqual(headers.get("Authorization"), "Bearer mytoken123")
        self.assertEqual(headers.get("X-API-Key"), "test_key")

        config = app.ScrapeRequest(
            url="https://example.com",
            auto_detect_pages=True,
            custom_cookies="sessionid=abc123xyz",
            custom_headers="Authorization: Bearer token",
            http_username="myuser",
            http_password="mypassword",
        )
        self.assertTrue(config.auto_detect_pages)
        self.assertEqual(config.custom_cookies, "sessionid=abc123xyz")
        self.assertEqual(config.http_username, "myuser")


class V2FeaturesTests(unittest.TestCase):
    def test_rotating_user_agents(self):
        self.assertGreaterEqual(len(app._USER_AGENTS), 5)
        ua = app._rotate_ua()
        self.assertIn(ua, app._USER_AGENTS)

    def test_content_hash(self):
        h1 = app.content_hash("<html> <body> <h1>Hello World</h1> </body> </html>")
        h2 = app.content_hash("<html>   <body>   <h1>Hello   World</h1>   </body>   </html>")
        h3 = app.content_hash("<html> <body> <h1>Different</h1> </body> </html>")
        self.assertEqual(h1, h2)
        self.assertNotEqual(h1, h3)

    def test_extract_opengraph_meta(self):
        raw_html = """
        <html><head>
          <meta property="og:title" content="Awesome Site">
          <meta property="og:description" content="Best site ever">
          <meta name="twitter:card" content="summary_large_image">
        </head><body></body></html>
        """
        soup = app.BeautifulSoup(raw_html, "lxml")
        meta = app.extract_opengraph_meta(soup, "https://example.com")
        self.assertEqual(meta.get("og_title"), "Awesome Site")
        self.assertEqual(meta.get("og_description"), "Best site ever")
        self.assertEqual(meta.get("twitter_card"), "summary_large_image")

    def test_extract_microdata_entities(self):
        raw_html = """
        <div itemscope itemtype="http://schema.org/LocalBusiness">
          <h2 itemprop="name">Acme Electrical LLC</h2>
          <span itemprop="telephone">212-555-0199</span>
          <a itemprop="email" href="mailto:info@acmeelectric.com">Email Us</a>
          <p itemprop="description">Commercial and residential electrical contracting services.</p>
        </div>
        """
        soup = app.BeautifulSoup(raw_html, "lxml")
        entities = app.extract_microdata_entities(soup, "https://example.com")
        self.assertEqual(len(entities), 1)
        entity = entities[0]
        self.assertEqual(entity["title"], "Acme Electrical LLC")
        self.assertIn("212-555-0199", entity["phone"])
        self.assertEqual(entity["email"], "info@acmeelectric.com")
        self.assertEqual(entity["contractor_type"], "Electrical Contractor")

    def test_extract_definition_lists(self):
        raw_html = """
        <dl>
          <dt>Company</dt><dd>Apex HVAC Systems</dd>
          <dt>Phone</dt><dd>(312) 555-9876</dd>
          <dt>Email</dt><dd>support@apexhvac.com</dd>
        </dl>
        """
        soup = app.BeautifulSoup(raw_html, "lxml")
        records = app.extract_definition_lists(soup, "https://example.com")
        self.assertEqual(len(records), 1)
        self.assertIn("support@apexhvac.com", records[0]["email"])
        self.assertEqual(records[0]["contractor_type"], "HVAC Contractor")

    def test_proxy_configuration(self):
        config = app.ScrapeRequest(
            url="https://example.com",
            proxy_url="http://user:pass@proxy.example.com:8080"
        )
        self.assertEqual(config.proxy_url, "http://user:pass@proxy.example.com:8080")

    def test_expanded_social_hosts(self):
        self.assertIn("github.com", app.SOCIAL_HOSTS)
        self.assertIn("reddit.com", app.SOCIAL_HOSTS)
        self.assertIn("threads.net", app.SOCIAL_HOSTS)
        self.assertIn("wa.me", app.SOCIAL_HOSTS)
        self.assertIn("discord.gg", app.SOCIAL_HOSTS)


if __name__ == "__main__":
    unittest.main()



