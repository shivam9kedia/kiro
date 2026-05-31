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
    targets = [
        ('WellnessWater-25-Pillar-Blog-Content.md',
         'WellnessWater-25-Pillar-Blog-Content.docx'),
        ('WellnessWater-Citations-and-Infographic-Briefs.md',
         'WellnessWater-Citations-and-Infographic-Briefs.docx'),
        ('WellnessWater-Publishing-Calendar-and-Checklist.md',
         'WellnessWater-Publishing-Calendar-and-Checklist.docx'),
        ('WellnessWater-15-Blog-Articles.md',
         'WellnessWater-15-Blog-Articles.docx'),
    ]
    for md_name, out_name in targets:
        build_docx.build_docx(os.path.join(here, md_name), os.path.join(here, out_name))
