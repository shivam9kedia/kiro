"""Slide XML generators for the Apollo Windows presentation."""
from .core import inches, SLIDE_WIDTH, SLIDE_HEIGHT


def _escape(text):
    """Escape XML special characters."""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")


def _textbox(id_num, name, left, top, width, height, paragraphs, anchor="t"):
    """Create a textbox shape with multiple paragraphs."""
    paras_xml = ""
    for p in paragraphs:
        text = _escape(p.get("text", ""))
        size = p.get("size", 1400)
        bold = "1" if p.get("bold") else "0"
        color = p.get("color", "1B2A4A")
        align = p.get("align", "l")
        spc_before = p.get("spc_before", 0)
        bullet = p.get("bullet", False)

        bullet_xml = ""
        if bullet:
            bullet_xml = f'<a:buFont typeface="Arial"/><a:buChar char="\u2022"/>'

        paras_xml += f'''      <a:p>
        <a:pPr algn="{align}">
          {bullet_xml}
          <a:spcBef><a:spcPts val="{spc_before}"/></a:spcBef>
        </a:pPr>
        <a:r>
          <a:rPr lang="en-US" sz="{size}" b="{bold}" dirty="0">
            <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
            <a:latin typeface="Segoe UI"/>
          </a:rPr>
          <a:t>{text}</a:t>
        </a:r>
      </a:p>
'''

    return f'''  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="{id_num}" name="{name}"/>
      <p:cNvSpPr txBox="1"/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="{left}" y="{top}"/>
        <a:ext cx="{width}" cy="{height}"/>
      </a:xfrm>
      <a:prstGeom prst="rect"/>
      <a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr wrap="square" anchor="{anchor}"/>
      <a:lstStyle/>
{paras_xml}    </p:txBody>
  </p:sp>
'''


def _rect(id_num, name, left, top, width, height, fill_color, border_color=None):
    """Create a filled rectangle shape."""
    border_xml = "<a:noFill/>" if not border_color else f'<a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill>'
    return f'''  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="{id_num}" name="{name}"/>
      <p:cNvSpPr/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="{left}" y="{top}"/>
        <a:ext cx="{width}" cy="{height}"/>
      </a:xfrm>
      <a:prstGeom prst="rect"/>
      <a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>
      <a:ln>{border_xml}</a:ln>
    </p:spPr>
  </p:sp>
'''


def _rounded_rect(id_num, name, left, top, width, height, fill_color, paragraphs=None):
    """Create a rounded rectangle with optional text."""
    text_xml = ""
    if paragraphs:
        paras_xml = ""
        for p in paragraphs:
            text = _escape(p.get("text", ""))
            size = p.get("size", 1200)
            bold = "1" if p.get("bold") else "0"
            color = p.get("color", "FFFFFF")
            align = p.get("align", "ctr")
            paras_xml += f'''      <a:p>
        <a:pPr algn="{align}"/>
        <a:r>
          <a:rPr lang="en-US" sz="{size}" b="{bold}">
            <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
            <a:latin typeface="Segoe UI"/>
          </a:rPr>
          <a:t>{text}</a:t>
        </a:r>
      </a:p>
'''
        text_xml = f'''    <p:txBody>
      <a:bodyPr wrap="square" anchor="ctr"/>
      <a:lstStyle/>
{paras_xml}    </p:txBody>'''

    return f'''  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="{id_num}" name="{name}"/>
      <p:cNvSpPr/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="{left}" y="{top}"/>
        <a:ext cx="{width}" cy="{height}"/>
      </a:xfrm>
      <a:prstGeom prst="roundRect"/>
      <a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>
      <a:ln><a:noFill/></a:ln>
    </p:spPr>
{text_xml}
  </p:sp>
'''


def wrap_slide(shapes_xml, bg_color="FFFFFF"):
    """Wrap shapes into a complete slide XML."""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
{shapes_xml}
    </p:spTree>
  </p:cSld>
</p:sld>'''
