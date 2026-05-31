"""
Builds the Wellness Water Company 25-Pillar Blog Content .docx
using the existing self-contained Markdown -> DOCX converter in build_docx.py.
"""

import os
import build_docx

# Customize the document core properties (title/author) for this deliverable.
build_docx.CORE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Wellness Water Company - 25 Pillar Blog Articles</dc:title>
  <dc:creator>Wellness Water Company - Content Team</dc:creator>
  <cp:lastModifiedBy>Wellness Water Company - Content Team</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-05-31T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-31T00:00:00Z</dcterms:modified>
</cp:coreProperties>'''

if __name__ == '__main__':
    here = os.path.dirname(os.path.abspath(__file__))
    md = os.path.join(here, 'WellnessWater-25-Pillar-Blog-Content.md')
    out = os.path.join(here, 'WellnessWater-25-Pillar-Blog-Content.docx')
    build_docx.build_docx(md, out)
