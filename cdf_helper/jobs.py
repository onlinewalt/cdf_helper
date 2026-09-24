"""Generation-job result types for CDF Helper.

A "generation job" (see CONTEXT.md) is one request to produce a 报关清单:
it parses sources, optionally enriches parts via AI, fills a template, and
returns a typed result that the web adapter renders.

This Module concentrates the *result* contract — the shape consumed by
``result.html`` — into one typed dataclass instead of letting it leak as a
stringly-typed dictionary assembled inline in ``webapp.do_generate`` (the
AI ``_worker`` closure) and re-stated by the synchronous ``_finish``
adapter. The deletion test now deletes ``JobResult``: complexity of the
template field names collapses to one place instead of being duplicated by
every render site.
"""

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class JobResult:
    """Typed result of a generation job; the single source of truth for the
    ``result.html`` template variables."""

    file_name: str
    vessel: str
    port: str
    date: str
    item_count: int
    warnings: List[str] = field(default_factory=list)
    ai_stats: Optional[dict] = None

    def to_template_dict(self) -> dict:
        """Flatten into the dict shape ``result.html`` expects.

        This is the seam between the typed domain Module and the Flask adapter:
        ``result.html`` should reference only these keys, never construct a new
        shape. Keeping the dict here means a template field rename is one edit,
        not a drift between the AI closure and the sync path.
        """
        return {
            "file_name": self.file_name,
            "vessel": self.vessel,
            "port": self.port,
            "date": self.date,
            "item_count": self.item_count,
            "warnings": self.warnings,
            "ai_stats": self.ai_stats,
        }
