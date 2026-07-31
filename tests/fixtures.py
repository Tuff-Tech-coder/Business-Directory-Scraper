"""HTML fixtures mirroring the structure of a real directory results page.

Kept as literals rather than saved .html files so the shape each test depends
on is visible next to the assertions. Every class name here matches a selector
used by directory_scraper.parse_listing / parse_results_page.
"""

# A complete organic listing with every field populated.
FULL_CARD = """
<div class="result">
  <a class="business-name"><span>Ace Plumbing &amp; Drain</span></a>
  <div class="phones phone primary">(512) 555-0142</div>
  <div class="street-address">1420 Commerce Blvd</div>
  <div class="locality">Austin, TX 78701</div>
  <a class="track-visit-website" href="https://aceplumbing.com/?utm_source=yp&amp;ref=123">Website</a>
  <div class="categories"><a>Plumbers</a><a>Drain Cleaning</a></div>
  <div class="result-rating four">
    <span class="count">(87)</span>
  </div>
</div>
"""

# Every optional element removed -- only the business name survives.
# This is the defensive-extraction case: missing markup must not raise.
MINIMAL_CARD = """
<div class="result">
  <a class="business-name"><span>Bare Minimum Plumbing</span></a>
</div>
"""

# No business name at all -- parse_listing must reject this outright.
NAMELESS_CARD = """
<div class="result">
  <div class="phones phone primary">(512) 555-0000</div>
  <div class="street-address">99 Nowhere St</div>
</div>
"""

# A paid placement. Carries the 'flash-endt' class alongside 'result'.
AD_CARD = """
<div class="result flash-endt">
  <a class="business-name"><span>Sponsored Plumbing Co</span></a>
  <div class="phones phone primary">(512) 555-9999</div>
</div>
"""

# Website link pointing at an internal redirect rather than an external site.
INTERNAL_REDIRECT_CARD = """
<div class="result">
  <a class="business-name"><span>Redirect Plumbing</span></a>
  <a class="track-visit-website" href="/redirect?to=somewhere">Website</a>
</div>
"""


def page(cards: str, next_href: str | None = "/search?page=2") -> str:
    """Wrap card markup in a results page, optionally with a next-page link."""
    nav = f'<a class="next" href="{next_href}">Next</a>' if next_href else ""
    return f"<html><body>{cards}{nav}</body></html>"
