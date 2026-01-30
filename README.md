# html-to-pptx

Converts HTML slide presentations into faithful `.pptx` files using `python-pptx`.

Fully generic: reads **all** colors, fonts, sizes, borders, and positions from the HTML and its `<style>` block. No hardcoded theme — works with any color scheme.

## Architecture overview

```
  input.html
      |
      |  python3 html_to_pptx.py
      |
      v
  1. PARSE ─────────────────────────────────────────────────────
      |
      |   lxml.html.fromstring()
      |       read raw HTML string into DOM tree
      |
      +──> DOM tree
      |
      |   _parse_stylesheet()
      |       regex-extract <style> block
      |       build { ".class": {prop: val} } dict
      |
      +──> stylesheet dict
      |
      v
  2. DISCOVER SLIDES ───────────────────────────────────────────
      |
      |   doc.cssselect("div.slide")
      |       find all slide containers
      |
      +──> slide_el[0], slide_el[1], ... slide_el[N]
      |
      v
  3. RESOLVE GLOBALS ───────────────────────────────────────────
      |
      |   _resolve_font(stylesheet)
      |       body / .slide font-family ──> "Georgia", "Arial", etc.
      |
      +──> font name
      |
      v
  4. RENDER EACH SLIDE ─────────────────────────────────────────
      |
      |   for each slide_el:
      |       SlideRenderer(pptx_slide, slide_el, stylesheet, font)
      |
      +──> per slide:
      |
      |   4a. _chrome()
      |       |   cssselect .top-bar ──> read height, background ──> _rect()
      |       |   cssselect .date-box ──> read pos, size, color ──> _rect() + _textbox()
      |       |   cssselect .main-title ──> read font-size, color ──> _textbox()
      |       |   cssselect .footer-bar ──> read height, background ──> _rect()
      |       |       .page-number ──> _textbox()
      |       |       .logo ──> _textbox()
      |       v
      |
      |   4b. _positioned_blocks()
      |       |   cssselect div[style] where position:absolute
      |       |       detect type by structure:
      |       |           has .section-header?
      |       |               has progress bar? ──> _planning()
      |       |               has <table>? ──> _render_table()
      |       |               else ──> _section_full()
      |       |           has <table> only? ──> _standalone_table()
      |       v
      |
      |   4c. _legend()
      |       |   detect by structure: absolute + bottom + colored <span>s
      |       |   _render_rich() ──> inline ● with colors
      |       v
      |
      |   4d. _links()
      |       |   cssselect a.link-text ──> color from stylesheet
      |       v
      |
      v
  5. SERIALIZE ─────────────────────────────────────────────────
      |
      |   prs.save("output.pptx")
      |
      v
  output.pptx
```

## Internal pipeline (detailed)

```
                          input.html
                              |
                              v
                    lxml.html.fromstring()
                              |
                    +---------+---------+
                    |                   |
                    v                   v
            <style> block          div.slide (x N)
                    |                   |
                    v                   |
          _parse_stylesheet()           |
                    |                   |
                    +----> SlideRenderer <----+
                                |
               +----------------+----------------+
               |                |                |
               v                v                v
           _chrome()    _positioned_blocks()   _legend() / _links()
               |                |
               |     +----------+----------+
               |     |          |          |
               v     v          v          v
          top-bar  section    table    planning
          date-box  chrome    (standalone)  |
          title     |                  +----+----+
          footer    v                  |         |
                _section_full()    progress   workload
                    |              bar        table
         +----------+----------+
         |          |          |
         v          v          v
      trend-box   table    _render_box_content()
                  (inline)       |
                          +------+------+
                          |      |      |
                          v      v      v
                        <p>    <ul>   <div>
                         |      |      |
                         v      v      v
                      _render_rich()  (recurse)
                              |
                              v
                     bold / color / circles
                              |
                              v
                         output.pptx
```

## Requirements

```
pip install -r requirements.txt
```

Python 3.8+

## Usage

```bash
# default: reads input.html, writes output.pptx
python html_to_pptx.py

# custom paths
python html_to_pptx.py my_slides.html my_presentation.pptx
```

## What it handles

- **Chrome**: top bar, date box, title, footer — dimensions, colors, fonts from CSS
- **Section boxes**: header separator, title, bordered content box — all from CSS
- **Tables**: column widths, header/cell backgrounds, font-size, text-align, font-weight, border color and style (solid/dashed) — from inline styles + stylesheet classes
- **Progress bars**: detected by structure (border-radius + height + background), not by color
- **Colored indicators**: circles (border-radius:50% spans) and mood bullets (● with color) rendered inline
- **Rich text**: bold, colored spans, nested strong/em preserved via recursive rendering
- **Lists**: `<ul>/<li>` with bullet character and color from CSS
- **Bullet items**: character and color from `.bullet-item::before` in stylesheet
- **Trend items**: font-size, weight, color, gap from `.trend-item` / `.trend-box` CSS
- **Links**: color and font-size from `.link-text` CSS
- **Legend**: detected by structure (absolute + bottom + colored circle spans), language-independent
- **Spacing**: margin-top, margin-bottom, margin-left, line-height respected
- **Font family**: resolved from `body` or `.slide` font-family in stylesheet
- **Dark/light detection**: luminance-based, not a hardcoded color list
