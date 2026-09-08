# ADR-0001: AI-enriched generation jobs run asynchronously behind a polling task

When the user enables "AI 智能填写" (use_ai), POST /generate returns a `processing.html`
page and the AI enrichment plus workbook generation run on a background thread; the
browser polls /status/<id> for live progress and /result/<id> for the final result
page. Without use_ai, /generate remains fully synchronous and returns result.html
directly.

- **Context**: AI enrichment can take seconds-to-minutes per batch (5s/batch × N).
  A synchronous request blocks the browser and risks proxy/Flask-dev-server timeouts
  with no user feedback; the result page must show a live progress log of batch calls.
- **Decision**: extract a generic task runner (cdf_helper/tasks.py: `submit`/`status`)
  that runs any work_fn on a daemon thread with an on_status log callback. The Flask
  view spawns one task on the AI path and renders `processing.html`; /status and
  /result become thin adapters over `tasks.status()`. The non-AI path is unchanged.
- **Consequence**: the public contract of POST /generate now has two outcomes
  (processing page vs result page) depending on use_ai. The result.html contract and
  the ai_stats/warnings payloads are preserved exactly across both paths.